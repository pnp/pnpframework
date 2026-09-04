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

        public static bool IsSnapshotOnly(ListFieldSnapshot field)
        {
            if (field == null || !field.ReadOnly || !field.Sealed
                || !SnapshotOnlyFields.TryGetValue(field.Id, out var reviewed))
            {
                return false;
            }

            return string.Equals(field.InternalName, reviewed.InternalName, StringComparison.OrdinalIgnoreCase)
                && string.Equals(field.TypeAsString, reviewed.TypeAsString, StringComparison.OrdinalIgnoreCase);
        }
    }
}
