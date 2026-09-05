using System;
using System.Collections.Generic;

namespace PnP.Framework.Migration.Lists.Fields
{
    /// <summary>
    /// Exact identities for SharePoint-owned cache or derived fields whose source
    /// schema and values are evidence, not portable migration input.
    /// </summary>
    internal static class ReviewedListRuntimeFieldCatalog
    {
        private static readonly IReadOnlyDictionary<Guid, (string InternalName, string TypeAsString)> SnapshotOnlyFields =
            new Dictionary<Guid, (string InternalName, string TypeAsString)>
            {
                [new Guid("8ea9462f-400b-428b-9537-ee4a245bc805")] = ("MediaServiceBillingMetadata", "Note"),
                [new Guid("b887b6b2-4dcf-34fc-98b1-d5a42c605755")] = ("MediaServiceFastMetadata", "Note"),
                [new Guid("617f8947-74b2-36bc-9f7e-21ded7029bb5")] = ("MediaServiceMetadata", "Note"),
                [new Guid("6e9ab27b-184d-411f-aedb-00386b5f89c2")] = ("MediaServiceObjectDetectorVersions", "Text"),
                [new Guid("dc412798-e1ac-4884-afa6-3cb4a10dcdda")] = ("MediaServiceSearchProperties", "Note"),
                [new Guid("d3c9caf7-044c-4c71-ae64-092981e54b33")] = ("SharedWithDetails", "Note"),
                [new Guid("ef991a83-108d-4407-8ee5-ccc0c3d836b9")] = ("SharedWithUsers", "UserMulti"),
                [new Guid("89a0fbbc-67b4-40bd-8c9a-386b325ab4ca")] = ("SharingHintHash", "Text")
            };

        private static readonly IReadOnlyDictionary<Guid, (string InternalName, string TypeAsString)> TaxonomyCacheFields =
            new Dictionary<Guid, (string InternalName, string TypeAsString)>
            {
                [new Guid("f3b0adf9-c1a2-4b02-920d-943fba4b3611")] = ("TaxCatchAll", "LookupMulti"),
                [new Guid("8f6b6dd8-9357-4019-8172-966fcd502ed2")] = ("TaxCatchAllLabel", "LookupMulti")
            };

        private static readonly IReadOnlyDictionary<Guid, (string InternalName, string TypeAsString)> ProtectedMetadataFields =
            new Dictionary<Guid, (string InternalName, string TypeAsString)>
            {
                [new Guid("97dfa283-6ba8-4e68-ab6f-846e53f6d381")] = ("_ip_UnifiedCompliancePolicyProperties", "Note"),
                [new Guid("a925f967-0e5c-443a-ab36-18749634253f")] = ("_ip_UnifiedCompliancePolicyUIAction", "Text")
            };

        public static bool IsSnapshotOnly(ListFieldSnapshot field)
        {
            if (field == null)
            {
                return false;
            }

            if (field.Hidden && field.Sealed
                && TaxonomyCacheFields.TryGetValue(field.Id, out var taxonomyCache))
            {
                return string.Equals(field.InternalName, taxonomyCache.InternalName, StringComparison.OrdinalIgnoreCase)
                    && string.Equals(field.TypeAsString, taxonomyCache.TypeAsString, StringComparison.OrdinalIgnoreCase);
            }

            if (ProtectedMetadataFields.TryGetValue(field.Id, out var protectedMetadata))
            {
                return string.Equals(field.InternalName, protectedMetadata.InternalName, StringComparison.OrdinalIgnoreCase)
                    && string.Equals(field.TypeAsString, protectedMetadata.TypeAsString, StringComparison.OrdinalIgnoreCase);
            }

            if (!field.ReadOnly || !field.Sealed
                || !SnapshotOnlyFields.TryGetValue(field.Id, out var reviewed))
            {
                return false;
            }

            return string.Equals(field.InternalName, reviewed.InternalName, StringComparison.OrdinalIgnoreCase)
                && string.Equals(field.TypeAsString, reviewed.TypeAsString, StringComparison.OrdinalIgnoreCase);
        }
    }
}
