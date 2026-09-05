using PnP.Framework.Migration.Pages.Publishing.Packaging;
using System;
using System.Text.RegularExpressions;

namespace PnP.Framework.Migration.Pages.Content
{
    internal static class PublishingPageContentStorageCanonicalizer
    {
        private static readonly Regex EncodedColon = new Regex(
            @"&(?:#0*58|#x0*3a|colon);",
            RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);

        public static string Canonicalize(string value)
        {
            // SharePoint can persist URI scheme separators as an HTML character
            // reference (for example https&#58;//). This is the same authored
            // character in HTML. Normalize only colon references rather than
            // decoding the entire fragment, which could turn escaped markup into
            // executable markup and hide a real storage difference.
            return EncodedColon.Replace(value ?? string.Empty, ":");
        }

        public static bool AreEquivalent(string expected, string actual)
        {
            return string.Equals(
                Canonicalize(expected),
                Canonicalize(actual),
                StringComparison.Ordinal);
        }

        public static string ComputeCanonicalSha256(string value)
        {
            return PublishingPageDigest.ComputeSha256(Canonicalize(value));
        }
    }
}
