using System;
using System.Linq;

namespace PnP.Framework.Migration.Pages.Publishing.Verification
{
    internal static class PublishingPageContentTypeIdentity
    {
        public static bool MatchesSiteContentType(string actual, string expectedSiteContentType)
        {
            if (string.IsNullOrWhiteSpace(actual) || string.IsNullOrWhiteSpace(expectedSiteContentType))
            {
                return false;
            }

            if (string.Equals(actual, expectedSiteContentType, StringComparison.OrdinalIgnoreCase))
            {
                return true;
            }

            // Content Type IDs are a hexadecimal path. A Pages library creates
            // a list-scoped descendant ID when a site Content Type is attached.
            // Require a complete additional segment rather than accepting an
            // arbitrary string prefix.
            if (!actual.StartsWith(expectedSiteContentType, StringComparison.OrdinalIgnoreCase))
            {
                return false;
            }
            var suffix = actual.Substring(expectedSiteContentType.Length);
            return suffix.Length >= 2
                && suffix.Length % 2 == 0
                && suffix.All(Uri.IsHexDigit);
        }
    }
}
