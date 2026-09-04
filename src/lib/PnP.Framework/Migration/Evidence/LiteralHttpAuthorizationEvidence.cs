using System;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text;

namespace PnP.Framework.Migration.Evidence
{
    /// <summary>
    /// Retains the minimum wire facts required to prove that a request returned a
    /// literal HTTP 401 or 403. Inferred access-denied messages and CSOM payload
    /// errors do not satisfy this contract.
    /// </summary>
    public sealed class LiteralHttpAuthorizationEvidence
    {
        public string Operation { get; set; }

        public string RequestUri { get; set; }

        public int HttpStatusCode { get; set; }

        public DateTimeOffset ObservedAtUtc { get; set; }

        public string EvidenceSha256 { get; set; }

        public static LiteralHttpAuthorizationEvidence Create(
            string operation,
            string requestUri,
            int httpStatusCode,
            DateTimeOffset observedAtUtc)
        {
            var evidence = new LiteralHttpAuthorizationEvidence
            {
                Operation = operation,
                RequestUri = requestUri,
                HttpStatusCode = httpStatusCode,
                ObservedAtUtc = observedAtUtc.ToUniversalTime()
            };
            evidence.EvidenceSha256 = ComputeSha256(evidence);
            Validate(evidence);
            return evidence;
        }

        public static void Validate(LiteralHttpAuthorizationEvidence evidence)
        {
            Uri requestUri;
            if (evidence == null
                || string.IsNullOrWhiteSpace(evidence.Operation)
                || !Uri.TryCreate(evidence.RequestUri, UriKind.Absolute, out requestUri)
                || requestUri.Scheme != Uri.UriSchemeHttp && requestUri.Scheme != Uri.UriSchemeHttps
                || evidence.HttpStatusCode != 401 && evidence.HttpStatusCode != 403
                || evidence.ObservedAtUtc == default(DateTimeOffset)
                || !IsSha256(evidence.EvidenceSha256)
                || !string.Equals(
                    evidence.EvidenceSha256,
                    ComputeSha256(evidence),
                    StringComparison.OrdinalIgnoreCase))
            {
                throw new InvalidDataException(
                    "Literal authorization evidence requires HTTP 401/403, operation, request URI, timestamp, and a matching SHA-256.");
            }
        }

        public static string ComputeSha256(LiteralHttpAuthorizationEvidence evidence)
        {
            if (evidence == null)
            {
                throw new ArgumentNullException(nameof(evidence));
            }

            var canonical = string.Join("\n", new[]
            {
                evidence.Operation?.Trim() ?? string.Empty,
                NormalizeUri(evidence.RequestUri),
                evidence.HttpStatusCode.ToString(CultureInfo.InvariantCulture),
                evidence.ObservedAtUtc.ToUniversalTime().ToString("O", CultureInfo.InvariantCulture)
            }) + "\n";
            using (var algorithm = SHA256.Create())
            {
                return string.Concat(algorithm.ComputeHash(Encoding.UTF8.GetBytes(canonical))
                    .Select(value => value.ToString("x2", CultureInfo.InvariantCulture)));
            }
        }

        private static string NormalizeUri(string value)
        {
            Uri uri;
            return Uri.TryCreate(value, UriKind.Absolute, out uri) ? uri.AbsoluteUri : value?.Trim() ?? string.Empty;
        }

        private static bool IsSha256(string value)
        {
            return !string.IsNullOrWhiteSpace(value)
                && value.Length == 64
                && value.All(character => character >= '0' && character <= '9'
                    || character >= 'a' && character <= 'f'
                    || character >= 'A' && character <= 'F');
        }
    }
}
