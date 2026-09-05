using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text;

namespace PnP.Framework.Migration.Topology
{
    /// <summary>
    /// Allocates a deterministic target URL segment while preserving the source segment
    /// whenever it is not occupied by another object.
    /// </summary>
    public static class TopologyTargetPathAllocator
    {
        /// <summary>
        /// Allocates a deterministic server-relative target path. The parent path is
        /// preserved exactly; only the occupied leaf receives a stable suffix.
        /// Occupied paths outside the same direct parent do not affect allocation.
        /// </summary>
        public static string AllocateServerRelativePath(
            string preferredPath,
            string stableSourceIdentity,
            IEnumerable<string> occupiedPaths,
            bool preserveFileExtension = false,
            int maximumSegmentLength = 128)
        {
            var normalizedPreferred = NormalizeServerRelativePath(preferredPath, nameof(preferredPath));
            var separator = normalizedPreferred.LastIndexOf('/');
            var parent = separator == 0 ? "/" : normalizedPreferred.Substring(0, separator);
            var preferredLeaf = normalizedPreferred.Substring(separator + 1);
            ValidateSegment(preferredLeaf, nameof(preferredPath));

            var occupiedLeaves = (occupiedPaths ?? Enumerable.Empty<string>())
                .Where(value => !string.IsNullOrWhiteSpace(value))
                .Select(value => NormalizeServerRelativePath(value, nameof(occupiedPaths)))
                .Where(value => string.Equals(ParentPath(value), parent, StringComparison.OrdinalIgnoreCase))
                .Select(LeafSegment)
                .ToArray();
            var allocatedLeaf = AllocateSegment(
                preferredLeaf,
                stableSourceIdentity,
                occupiedLeaves,
                preserveFileExtension,
                maximumSegmentLength);
            return parent == "/" ? "/" + allocatedLeaf : parent + "/" + allocatedLeaf;
        }

        public static string AllocateSegment(
            string preferredSegment,
            string stableSourceIdentity,
            IEnumerable<string> occupiedSegments,
            bool preserveFileExtension = false,
            int maximumLength = 128)
        {
            ValidateSegment(preferredSegment, nameof(preferredSegment));
            if (string.IsNullOrWhiteSpace(stableSourceIdentity))
            {
                throw new ArgumentException("A stable source identity is required.", nameof(stableSourceIdentity));
            }
            if (maximumLength < 16)
            {
                throw new ArgumentOutOfRangeException(nameof(maximumLength), maximumLength, "The maximum segment length must be at least 16 characters.");
            }

            var occupied = new HashSet<string>(
                occupiedSegments ?? Enumerable.Empty<string>(),
                StringComparer.OrdinalIgnoreCase);
            if (!occupied.Contains(preferredSegment))
            {
                return preferredSegment;
            }

            var digest = StableDigest(stableSourceIdentity);
            for (var digestLength = 8; digestLength <= digest.Length; digestLength += 4)
            {
                var suffix = "-pnp-" + digest.Substring(0, digestLength);
                var candidate = AddSuffix(preferredSegment, suffix, preserveFileExtension, maximumLength);
                if (!occupied.Contains(candidate))
                {
                    return candidate;
                }
            }

            throw new InvalidOperationException("No deterministic target segment remained after exhausting the stable source digest.");
        }

        private static string AddSuffix(string segment, string suffix, bool preserveFileExtension, int maximumLength)
        {
            var extension = preserveFileExtension ? Path.GetExtension(segment) : string.Empty;
            var stem = extension.Length == 0 ? segment : segment.Substring(0, segment.Length - extension.Length);
            var maximumStemLength = maximumLength - suffix.Length - extension.Length;
            if (maximumStemLength < 1)
            {
                throw new InvalidOperationException("The target segment length does not leave room for a deterministic collision suffix.");
            }
            if (stem.Length > maximumStemLength)
            {
                stem = stem.Substring(0, maximumStemLength).TrimEnd(' ', '.', '-');
            }
            if (stem.Length == 0)
            {
                stem = "item";
            }
            return stem + suffix + extension;
        }

        private static string StableDigest(string value)
        {
            using (var algorithm = SHA256.Create())
            {
                return BitConverter.ToString(algorithm.ComputeHash(Encoding.UTF8.GetBytes(value)))
                    .Replace("-", string.Empty)
                    .ToLowerInvariant();
            }
        }

        private static void ValidateSegment(string value, string parameterName)
        {
            if (string.IsNullOrWhiteSpace(value)
                || value == "."
                || value == ".."
                || value.IndexOf('/') >= 0
                || value.IndexOf('\\') >= 0)
            {
                throw new ArgumentException("A safe single URL segment is required.", parameterName);
            }
        }

        private static string NormalizeServerRelativePath(string value, string parameterName)
        {
            if (string.IsNullOrWhiteSpace(value))
            {
                throw new ArgumentException("A server-relative path is required.", parameterName);
            }

            var normalized = Uri.UnescapeDataString(value.Trim()).Replace('\\', '/');
            if (!normalized.StartsWith("/", StringComparison.Ordinal) || normalized == "/")
            {
                throw new ArgumentException("A non-root server-relative path is required.", parameterName);
            }
            normalized = normalized.TrimEnd('/');
            if (normalized.Split('/').Any(segment => segment == "." || segment == ".."))
            {
                throw new ArgumentException("Relative traversal is not allowed.", parameterName);
            }
            return normalized;
        }

        private static string ParentPath(string path)
        {
            var separator = path.LastIndexOf('/');
            return separator == 0 ? "/" : path.Substring(0, separator);
        }

        private static string LeafSegment(string path)
        {
            return path.Substring(path.LastIndexOf('/') + 1);
        }
    }
}
