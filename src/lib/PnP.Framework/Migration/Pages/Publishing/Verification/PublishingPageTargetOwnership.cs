using PnP.Framework.Migration.Pages;
using System;
using System.Collections.Generic;

namespace PnP.Framework.Migration.Pages.Publishing.Verification
{
    internal static class PublishingPageTargetOwnership
    {
        public const string OriginalIdentifierPropertyName = "pnp_reserved_page_original_identifier";

        public const string SourceSnapshotDigestPropertyName = "pnp_reserved_page_source_snapshot_digest";

        public const string PlanDigestPropertyName = "pnp_reserved_page_migration_digest";

        public static string OriginalIdentifier(PageIdentity source)
        {
            if (source == null)
            {
                throw new ArgumentNullException(nameof(source));
            }
            if (source.SiteId == Guid.Empty || source.WebId == Guid.Empty || source.FileUniqueId == Guid.Empty)
            {
                throw new ArgumentException("A source Site, Web, and file identity are required.", nameof(source));
            }
            return "urn:pnp:spo-page:v1:" + source.SiteId.ToString("D") + ":"
                + source.WebId.ToString("D") + ":" + source.FileUniqueId.ToString("D");
        }

        public static bool MatchesApprovedPlan(
            IDictionary<string, object> properties,
            string originalIdentifier,
            string sourceSnapshotDigest,
            string planDigest)
        {
            return properties != null
                && Matches(properties, OriginalIdentifierPropertyName, originalIdentifier, StringComparison.Ordinal)
                && Matches(properties, SourceSnapshotDigestPropertyName, sourceSnapshotDigest, StringComparison.OrdinalIgnoreCase)
                && Matches(properties, PlanDigestPropertyName, planDigest, StringComparison.OrdinalIgnoreCase);
        }

        private static bool Matches(
            IDictionary<string, object> properties,
            string key,
            string expected,
            StringComparison comparison)
        {
            object value;
            return !string.IsNullOrWhiteSpace(expected)
                && properties.TryGetValue(key, out value)
                && string.Equals(value?.ToString(), expected, comparison);
        }
    }
}
