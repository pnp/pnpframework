using System;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text;

namespace PnP.Framework.Migration.Evidence
{
    /// <summary>
    /// Retains the literal SharePoint wire response that proves content is stored
    /// in Microsoft 365 Archive and must be reactivated before its bytes can be read.
    /// This is a recoverable source-content state, not an authorization blocker.
    /// </summary>
    public sealed class LiteralHttpArchivedContentEvidence
    {
        public string Operation { get; set; }

        public string RequestUri { get; set; }

        public int HttpStatusCode { get; set; }

        public string ErrorCode { get; set; }

        public string InnerErrorCode { get; set; }

        public string Message { get; set; }

        public DateTimeOffset ObservedAtUtc { get; set; }

        public string EvidenceSha256 { get; set; }

        public static LiteralHttpArchivedContentEvidence Create(
            string operation,
            string requestUri,
            int httpStatusCode,
            string errorCode,
            string innerErrorCode,
            string message,
            DateTimeOffset observedAtUtc)
        {
            var evidence = new LiteralHttpArchivedContentEvidence
            {
                Operation = operation,
                RequestUri = requestUri,
                HttpStatusCode = httpStatusCode,
                ErrorCode = errorCode,
                InnerErrorCode = innerErrorCode,
                Message = message,
                ObservedAtUtc = observedAtUtc.ToUniversalTime()
            };
            evidence.EvidenceSha256 = ComputeSha256(evidence);
            Validate(evidence);
            return evidence;
        }

        public static void Validate(LiteralHttpArchivedContentEvidence evidence)
        {
            Uri requestUri;
            if (evidence == null
                || string.IsNullOrWhiteSpace(evidence.Operation)
                || !Uri.TryCreate(evidence.RequestUri, UriKind.Absolute, out requestUri)
                || requestUri.Scheme != Uri.UriSchemeHttp && requestUri.Scheme != Uri.UriSchemeHttps
                || evidence.HttpStatusCode != 423
                || !string.Equals(evidence.ErrorCode, "locked", StringComparison.OrdinalIgnoreCase)
                || !string.Equals(evidence.InnerErrorCode, "contentArchived", StringComparison.OrdinalIgnoreCase)
                || string.IsNullOrWhiteSpace(evidence.Message)
                || evidence.ObservedAtUtc == default(DateTimeOffset)
                || !IsSha256(evidence.EvidenceSha256)
                || !string.Equals(
                    evidence.EvidenceSha256,
                    ComputeSha256(evidence),
                    StringComparison.OrdinalIgnoreCase))
            {
                throw new InvalidDataException(
                    "Archived-content evidence requires literal HTTP 423 locked/contentArchived, operation, request URI, message, timestamp, and a matching SHA-256.");
            }
        }

        public static string ComputeSha256(LiteralHttpArchivedContentEvidence evidence)
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
                evidence.ErrorCode?.Trim() ?? string.Empty,
                evidence.InnerErrorCode?.Trim() ?? string.Empty,
                evidence.Message?.Trim() ?? string.Empty,
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
