using Microsoft.SharePoint.Client;
using PnP.Framework.Migration.Pages.Content;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;

namespace PnP.Framework.Migration.Pages.Fields
{
    internal static class PageFieldWriter
    {
        public static List<PageFieldImportResult> Apply(
            ClientContext context,
            ListItem targetItem,
            IEnumerable<PageFieldValueSnapshot> fields,
            IEnumerable<PageFieldAction> actions,
            IEnumerable<PageTextReplacement> replacements,
            ICollection<string> warnings)
        {
            var fieldByName = fields.ToDictionary(field => field.InternalName, StringComparer.OrdinalIgnoreCase);
            var results = new List<PageFieldImportResult>();
            foreach (var action in actions)
            {
                var result = new PageFieldImportResult
                {
                    InternalName = action.SourceInternalName,
                    PlannedDisposition = action.Disposition,
                    Attempted = action.Disposition == PageFieldDisposition.Apply,
                    Succeeded = false,
                    Message = action.Reason
                };
                results.Add(result);
                if (!result.Attempted)
                {
                    continue;
                }

                if (!fieldByName.TryGetValue(action.SourceInternalName, out var field))
                {
                    result.Message = "The planned source field is missing from the sealed snapshot.";
                    warnings.Add($"Planned field '{action.SourceInternalName}' is missing from the sealed snapshot.");
                    continue;
                }

                try
                {
                    SetValue(targetItem, action.TargetInternalName, field, replacements);
                    targetItem.Update();
                    context.ExecuteQueryRetry();
                    result.Succeeded = true;
                    result.Message = "Applied successfully.";
                }
                catch (Exception exception)
                {
                    result.Message = exception.Message;
                    warnings.Add($"Field '{action.SourceInternalName}' could not be applied: {exception.Message}");
                }
            }

            return results;
        }

        private static void SetValue(
            ListItem targetItem,
            string targetInternalName,
            PageFieldValueSnapshot field,
            IEnumerable<PageTextReplacement> replacements)
        {
            switch (field.Kind)
            {
                case PageFieldValueKind.String:
                    targetItem[targetInternalName] = PageTextTransformer.Rewrite(field.Value, replacements);
                    break;
                case PageFieldValueKind.StringCollection:
                    targetItem[targetInternalName] = field.StringValues.ToArray();
                    break;
                case PageFieldValueKind.Boolean:
                    targetItem[targetInternalName] = bool.Parse(field.Value);
                    break;
                case PageFieldValueKind.Number:
                    targetItem[targetInternalName] = double.Parse(field.Value, NumberStyles.Any, CultureInfo.InvariantCulture);
                    break;
                case PageFieldValueKind.DateTime:
                    targetItem[targetInternalName] = DateTime.Parse(field.Value, CultureInfo.InvariantCulture, DateTimeStyles.RoundtripKind);
                    break;
                case PageFieldValueKind.Guid:
                    targetItem[targetInternalName] = Guid.Parse(field.Value);
                    break;
                case PageFieldValueKind.Url:
                    targetItem[targetInternalName] = new FieldUrlValue
                    {
                        Url = PageTextTransformer.Rewrite(field.UrlValue?.Url, replacements),
                        Description = field.UrlValue?.Description
                    };
                    break;
                default:
                    throw new NotSupportedException($"Field value kind '{field.Kind}' is not importable.");
            }
        }
    }
}
