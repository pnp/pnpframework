using PnP.Framework.Migration.Evidence;
using System;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Security.Cryptography;

namespace PnP.Framework.Migration.Packaging
{
    /// <summary>
    /// Stores migration payloads by SHA-256 under a local directory.
    /// </summary>
    public sealed class DirectoryMigrationArtifactStore : IMigrationArtifactStore
    {
        private readonly string rootDirectory;

        public DirectoryMigrationArtifactStore(string rootDirectory)
        {
            if (string.IsNullOrWhiteSpace(rootDirectory))
            {
                throw new ArgumentException("An artifact-store directory is required.", nameof(rootDirectory));
            }

            this.rootDirectory = Path.GetFullPath(rootDirectory);
            Directory.CreateDirectory(this.rootDirectory);
        }

        public string RootDirectory => rootDirectory;

        public bool Contains(string sha256)
        {
            return File.Exists(GetArtifactPath(sha256));
        }

        public Stream OpenRead(string sha256)
        {
            var path = GetArtifactPath(sha256);
            if (!File.Exists(path))
            {
                throw new FileNotFoundException($"Migration artifact '{sha256}' was not found.", path);
            }

            return new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read, 81920, FileOptions.SequentialScan);
        }

        public ArtifactReference Put(Stream content, string mediaType = null, string originalName = null)
        {
            if (content == null)
            {
                throw new ArgumentNullException(nameof(content));
            }

            var incomingPath = Path.Combine(rootDirectory, $".incoming-{Guid.NewGuid():N}");
            string digest = null;
            long length = 0;
            try
            {
                using (var algorithm = SHA256.Create())
                using (var destination = new FileStream(
                    incomingPath,
                    FileMode.CreateNew,
                    FileAccess.Write,
                    FileShare.None,
                    81920,
                    FileOptions.SequentialScan))
                {
                    var buffer = new byte[81920];
                    int read;
                    while ((read = content.Read(buffer, 0, buffer.Length)) > 0)
                    {
                        destination.Write(buffer, 0, read);
                        algorithm.TransformBlock(buffer, 0, read, null, 0);
                        length += read;
                    }

                    algorithm.TransformFinalBlock(Array.Empty<byte>(), 0, 0);
                    destination.Flush(true);
                    digest = string.Concat(algorithm.Hash.Select(value => value.ToString("x2", CultureInfo.InvariantCulture)));
                }

                var artifactPath = GetArtifactPath(digest);
                Directory.CreateDirectory(Path.GetDirectoryName(artifactPath));
                if (File.Exists(artifactPath))
                {
                    VerifyExisting(artifactPath, digest, length);
                }
                else
                {
                    try
                    {
                        File.Move(incomingPath, artifactPath);
                    }
                    catch (IOException) when (File.Exists(artifactPath))
                    {
                        VerifyExisting(artifactPath, digest, length);
                    }
                }

                return new ArtifactReference
                {
                    Sha256 = digest,
                    Length = length,
                    MediaType = mediaType,
                    OriginalName = originalName,
                    Availability = EvidenceAvailability.Captured
                };
            }
            finally
            {
                if (File.Exists(incomingPath))
                {
                    File.Delete(incomingPath);
                }
            }
        }

        private string GetArtifactPath(string sha256)
        {
            var digest = NormalizeDigest(sha256);
            var path = Path.GetFullPath(Path.Combine(rootDirectory, digest.Substring(0, 2), digest));
            var rootPrefix = rootDirectory.TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar)
                + Path.DirectorySeparatorChar;
            if (!path.StartsWith(rootPrefix, StringComparison.OrdinalIgnoreCase))
            {
                throw new InvalidDataException("The artifact digest resolved outside the configured store.");
            }

            return path;
        }

        private static string NormalizeDigest(string sha256)
        {
            var digest = (sha256 ?? string.Empty).Trim().ToLowerInvariant();
            if (digest.Length != 64 || digest.Any(value => !Uri.IsHexDigit(value)))
            {
                throw new ArgumentException("A 64-character SHA-256 digest is required.", nameof(sha256));
            }

            return digest;
        }

        private static void VerifyExisting(string path, string expectedDigest, long expectedLength)
        {
            var info = new FileInfo(path);
            if (info.Length != expectedLength)
            {
                throw new InvalidDataException($"Existing artifact '{expectedDigest}' has an unexpected length.");
            }

            string actualDigest;
            using (var content = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read, 81920, FileOptions.SequentialScan))
            {
                actualDigest = MigrationDigest.ComputeSha256(content);
            }

            if (!string.Equals(actualDigest, expectedDigest, StringComparison.OrdinalIgnoreCase))
            {
                throw new InvalidDataException($"Existing artifact '{expectedDigest}' is corrupt.");
            }
        }
    }
}
