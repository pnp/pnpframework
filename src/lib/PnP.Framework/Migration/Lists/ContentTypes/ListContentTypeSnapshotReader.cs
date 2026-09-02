using Microsoft.SharePoint.Client;
using System;
using System.Collections.Generic;
using System.Linq;

namespace PnP.Framework.Migration.Lists.ContentTypes
{
    internal static class ListContentTypeSnapshotReader
    {
        public static IList<ListContentTypeSnapshot> Read(
            ContentTypeCollection contentTypes,
            ICollection<string> diagnostics)
        {
            return contentTypes.AsEnumerable().Select(contentType =>
            {
                var fieldLinks = contentType.FieldLinks.Select(link => new ListContentTypeFieldLinkSnapshot
                {
                    FieldId = link.Id,
                    InternalName = link.Name,
                    DisplayName = link.DisplayName,
                    Required = link.Required,
                    Hidden = link.Hidden,
                    ReadOnly = link.ReadOnly
                }).ToArray();
                var canonicalLinks = new List<ListContentTypeFieldLinkSnapshot>();
                foreach (var group in fieldLinks.GroupBy(value => value.FieldId).OrderBy(value => value.Key))
                {
                    var variants = group
                        .OrderBy(FieldLinkIdentity, StringComparer.Ordinal)
                        .ToArray();
                    var identities = variants.Select(FieldLinkIdentity).Distinct(StringComparer.Ordinal).ToArray();
                    if (variants.Length > 1)
                    {
                        var prefix = identities.Length == 1
                            ? "DuplicateListContentTypeFieldLink"
                            : "ConflictingListContentTypeFieldLink";
                        diagnostics?.Add(prefix + ": List content type '" + contentType.Id.StringValue
                            + "' returned " + variants.Length + " rows for field link '" + group.Key.ToString("D")
                            + "'. Retained='" + identities[0] + "'; observed='" + string.Join(" || ", identities) + "'.");
                    }
                    canonicalLinks.Add(variants[0]);
                }

                return new ListContentTypeSnapshot
                {
                    Id = contentType.Id.StringValue,
                    Name = contentType.Name,
                    Description = contentType.Description ?? string.Empty,
                    Group = contentType.Group ?? string.Empty,
                    ParentId = contentType.Parent == null ? null : contentType.Parent.Id.StringValue,
                    Hidden = contentType.Hidden,
                    ReadOnly = contentType.ReadOnly,
                    Sealed = contentType.Sealed,
                    FieldLinks = canonicalLinks
                };
            }).OrderBy(value => value.Id, StringComparer.OrdinalIgnoreCase).ToList();
        }

        private static string FieldLinkIdentity(ListContentTypeFieldLinkSnapshot value)
        {
            return (value.InternalName ?? string.Empty) + "\u001f"
                + (value.DisplayName ?? string.Empty) + "\u001f"
                + value.Required + "\u001f" + value.Hidden + "\u001f" + value.ReadOnly;
        }
    }
}
