using PnP.Framework.Utilities;
using System;
using System.Linq;

namespace PnP.Framework.Migration.Pages
{
    internal static class PagePath
    {
        public static string Normalize(string webServerRelativeUrl, string value, string defaultLibrary)
        {
            var candidate = value.Trim();
            if (Uri.TryCreate(candidate, UriKind.Absolute, out var absolute))
            {
                candidate = Uri.UnescapeDataString(absolute.AbsolutePath);
            }
            else
            {
                candidate = Uri.UnescapeDataString(candidate.Split(new[] { '?', '#' }, 2)[0]).Replace('\\', '/');
            }

            if (!candidate.StartsWith("/", StringComparison.Ordinal))
            {
                if (!candidate.Contains("/"))
                {
                    candidate = UrlUtility.Combine(defaultLibrary, candidate);
                }

                candidate = UrlUtility.Combine(webServerRelativeUrl, candidate);
            }

            if (!candidate.EndsWith(".aspx", StringComparison.OrdinalIgnoreCase))
            {
                candidate += ".aspx";
            }

            return candidate;
        }

        public static bool IsWithin(string candidate, string root)
        {
            var normalizedCandidate = (candidate ?? string.Empty).TrimEnd('/');
            var normalizedRoot = (root ?? string.Empty).TrimEnd('/');
            return string.Equals(normalizedCandidate, normalizedRoot, StringComparison.OrdinalIgnoreCase)
                || normalizedCandidate.StartsWith(normalizedRoot + "/", StringComparison.OrdinalIgnoreCase);
        }

        public static bool UriEquals(string left, string right)
        {
            return string.Equals(left?.TrimEnd('/'), right?.TrimEnd('/'), StringComparison.OrdinalIgnoreCase);
        }

        public static string GetDirectoryName(string serverRelativeUrl)
        {
            var separator = serverRelativeUrl.LastIndexOf('/');
            return separator <= 0 ? "/" : serverRelativeUrl.Substring(0, separator);
        }

        public static string GetFileName(string serverRelativeUrl)
        {
            var separator = serverRelativeUrl.LastIndexOf('/');
            return separator < 0 ? serverRelativeUrl : serverRelativeUrl.Substring(separator + 1);
        }

        public static string Encode(string decodedPath)
        {
            return string.Join("/", decodedPath.Split('/').Select(Uri.EscapeDataString));
        }
    }
}
