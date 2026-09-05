using Microsoft.SharePoint.Client;
using PnP.Framework.Migration.Pages.Capture;
using PnP.Framework.Migration.Pages.Fields.Taxonomy;
using System;
using System.Collections.Generic;
using System.Linq;

namespace PnP.Framework.Migration.Pages.Fields
{
    internal static class PageFieldSnapshotReader
    {
        public static List<PageFieldValueSnapshot> Read(
            ClientContext context,
            ListItem item,
            ICollection<string> warnings)
        {
            var parentList = item.ParentList;
            context.Load(parentList.Fields, fields => fields.Include(
                field => field.Id,
                field => field.InternalName,
                field => field.Title,
                field => field.TypeAsString,
                field => field.SchemaXml,
                field => field.ReadOnlyField,
                field => field.Hidden,
                field => field.Required));
            context.ExecuteQueryRetry();

            var result = new List<PageFieldValueSnapshot>();
            foreach (var field in parentList.Fields.OrderBy(value => value.InternalName, StringComparer.Ordinal))
            {
                if (!item.FieldValues.TryGetValue(field.InternalName, out var value))
                {
                    result.Add(PageFieldValueSerializer.Create(field, PageCaptureStatus.NotReturned, PageFieldValueKind.Null));
                    continue;
                }

                try
                {
                    var snapshot = PageFieldValueSerializer.Serialize(field, value);
                    result.Add(snapshot);
                    if (snapshot.CaptureStatus == PageCaptureStatus.CapturedWithLimitations)
                    {
                        warnings.Add($"Field '{field.InternalName}' has value type '{snapshot.RawType}' that was captured as recovery evidence only.");
                    }
                }
                catch (Exception exception)
                {
                    var snapshot = PageFieldValueSerializer.Create(field, PageCaptureStatus.Failed, PageFieldValueKind.Unsupported);
                    snapshot.HasValue = value != null;
                    snapshot.RawType = value?.GetType().FullName;
                    snapshot.RawValue = PageFieldValueSerializer.SafeConvertToString(value);
                    snapshot.Diagnostics.Add(exception.Message);
                    result.Add(snapshot);
                    warnings.Add($"Field '{field.InternalName}' could not be fully serialized and remains in the snapshot with diagnostics: {exception.Message}");
                }
            }

            PageTaxonomyRelationshipSnapshotReader.Enrich(context, item, result, warnings);
            return result;
        }
    }
}
