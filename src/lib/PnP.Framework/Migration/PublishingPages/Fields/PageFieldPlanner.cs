using Microsoft.SharePoint.Client;
using PnP.Framework.Migration.PublishingPages.Capture;
using PnP.Framework.Migration.PublishingPages.Planning;
using System;
using System.Collections.Generic;
using System.Linq;

namespace PnP.Framework.Migration.PublishingPages.Fields
{
    internal static class PageFieldPlanner
    {
        public static List<PageFieldAction> BuildActions(
            ClientContext targetContext,
            IEnumerable<PageFieldValueSnapshot> fields,
            ISet<string> handledFieldNames,
            ISet<string> recognizedFieldNames,
            PublishingPagePlanningOptions options,
            ICollection<string> warnings)
        {
            var pages = targetContext.Web.GetPagesLibrary();
            var sourceFields = fields.OrderBy(field => field.InternalName, StringComparer.Ordinal).ToArray();
            if (pages == null)
            {
                return sourceFields.Select(field => new PageFieldAction
                {
                    SourceInternalName = field.InternalName,
                    TargetInternalName = field.InternalName,
                    Disposition = PageFieldDisposition.Block,
                    Reason = "The target publishing Pages library is unavailable."
                }).ToList();
            }

            targetContext.Load(pages.Fields, values => values.Include(
                field => field.InternalName,
                field => field.TypeAsString,
                field => field.ReadOnlyField));
            targetContext.ExecuteQueryRetry();
            var targetFields = pages.Fields.ToDictionary(field => field.InternalName, StringComparer.OrdinalIgnoreCase);
            var result = new List<PageFieldAction>();
            foreach (var sourceField in sourceFields)
            {
                var action = new PageFieldAction
                {
                    SourceInternalName = sourceField.InternalName,
                    TargetInternalName = sourceField.InternalName
                };
                result.Add(action);
                if (handledFieldNames.Contains(sourceField.InternalName))
                {
                    action.Disposition = PageFieldDisposition.AlreadyHandled;
                    action.Reason = "The page creation workflow handles this field explicitly.";
                    continue;
                }

                if (sourceField.CaptureStatus == PageCaptureStatus.Failed
                    || sourceField.CaptureStatus == PageCaptureStatus.NotReturned)
                {
                    action.Disposition = PageFieldDisposition.CaptureUnavailable;
                    action.Reason = "The field definition is preserved, but no restorable value was captured.";
                    continue;
                }

                if (!sourceField.HasValue || sourceField.Kind == PageFieldValueKind.Null)
                {
                    action.Disposition = PageFieldDisposition.SkipEmpty;
                    action.Reason = "The source item has no value for this field.";
                    continue;
                }

                if (!recognizedFieldNames.Contains(sourceField.InternalName))
                {
                    action.Disposition = PageFieldDisposition.EvidenceOnly;
                    action.Reason = "The field is fully retained in the snapshot, but this importer does not recognize it yet.";
                    continue;
                }

                if (sourceField.ReadOnly)
                {
                    action.Disposition = PageFieldDisposition.SkipReadOnly;
                    action.Reason = "The source field is read-only.";
                    continue;
                }

                if (string.Equals(sourceField.TypeAsString, "Calculated", StringComparison.OrdinalIgnoreCase))
                {
                    action.Disposition = PageFieldDisposition.SkipCalculated;
                    action.Reason = "Calculated fields are recomputed by SharePoint.";
                    continue;
                }

                if (!targetFields.TryGetValue(sourceField.InternalName, out var targetField))
                {
                    action.Disposition = PageFieldDisposition.TargetFieldMissing;
                    action.Reason = "The recognized field is not present in the target Pages library.";
                    warnings.Add($"Recognized field '{sourceField.InternalName}' is absent from the target Pages library and will not be applied.");
                    continue;
                }

                action.TargetInternalName = targetField.InternalName;
                action.TargetTypeAsString = targetField.TypeAsString;
                if (targetField.ReadOnlyField)
                {
                    action.Disposition = PageFieldDisposition.SkipReadOnly;
                    action.Reason = "The target field is read-only.";
                    continue;
                }

                if (!string.Equals(sourceField.TypeAsString, targetField.TypeAsString, StringComparison.OrdinalIgnoreCase))
                {
                    action.Disposition = PageFieldDisposition.TargetTypeMismatch;
                    action.Reason = $"Source type '{sourceField.TypeAsString}' does not match target type '{targetField.TypeAsString}'.";
                    warnings.Add($"Recognized field '{sourceField.InternalName}' has a target type mismatch and will not be applied.");
                    continue;
                }

                if (RequiresIdentityMapping(sourceField.Kind))
                {
                    action.Disposition = PageFieldDisposition.RequiresMapping;
                    action.Reason = "The value is captured, but its source identity must be mapped before it can be safely applied to another site.";
                    if (!IsTaxonomy(sourceField.Kind) || !options.BlockOnManagedMetadata)
                    {
                        warnings.Add($"Field '{sourceField.InternalName}' requires an identity or term mapping and remains evidence-only.");
                    }

                    continue;
                }

                if (!IsImportableKind(sourceField.Kind))
                {
                    action.Disposition = PageFieldDisposition.EvidenceOnly;
                    action.Reason = $"No importer is registered for value kind '{sourceField.Kind}'.";
                    continue;
                }

                action.Disposition = PageFieldDisposition.Apply;
                action.Reason = "The field is recognized, writable, type-compatible, and has a supported captured value.";
            }

            return result;
        }

        public static bool IsImportableKind(PageFieldValueKind kind)
        {
            return kind == PageFieldValueKind.String
                || kind == PageFieldValueKind.StringCollection
                || kind == PageFieldValueKind.Boolean
                || kind == PageFieldValueKind.Number
                || kind == PageFieldValueKind.DateTime
                || kind == PageFieldValueKind.Guid
                || kind == PageFieldValueKind.Url;
        }

        private static bool RequiresIdentityMapping(PageFieldValueKind kind)
        {
            return IsTaxonomy(kind)
                || kind == PageFieldValueKind.User
                || kind == PageFieldValueKind.UserCollection
                || kind == PageFieldValueKind.Lookup
                || kind == PageFieldValueKind.LookupCollection;
        }

        private static bool IsTaxonomy(PageFieldValueKind kind)
        {
            return kind == PageFieldValueKind.Taxonomy || kind == PageFieldValueKind.TaxonomyCollection;
        }
    }
}
