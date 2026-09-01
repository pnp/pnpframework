using Microsoft.SharePoint.Client;
using System;
using System.Collections.Generic;
using System.Linq;

namespace PnP.Framework.Migration.Lists.ContentTypes
{
    internal static class ListContentTypeSnapshotReader
    {
        public static IList<ListContentTypeSnapshot> Read(ContentTypeCollection contentTypes)
        {
            return contentTypes.Select(contentType => new ListContentTypeSnapshot
            {
                Id = contentType.Id.StringValue,
                Name = contentType.Name,
                Description = contentType.Description ?? string.Empty,
                Group = contentType.Group ?? string.Empty,
                ParentId = contentType.Parent == null ? null : contentType.Parent.Id.StringValue,
                Hidden = contentType.Hidden,
                ReadOnly = contentType.ReadOnly,
                Sealed = contentType.Sealed,
                FieldLinks = contentType.FieldLinks.Select(link => new ListContentTypeFieldLinkSnapshot
                {
                    FieldId = link.Id,
                    InternalName = link.Name,
                    DisplayName = link.DisplayName,
                    Required = link.Required,
                    Hidden = link.Hidden,
                    ReadOnly = link.ReadOnly
                }).ToList()
            }).OrderBy(value => value.Id, StringComparer.OrdinalIgnoreCase).ToList();
        }
    }
}
