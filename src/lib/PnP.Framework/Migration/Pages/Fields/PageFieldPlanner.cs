using Microsoft.SharePoint.Client;
using PnP.Framework.Migration.Pages.Capture;
using PnP.Framework.Migration.Pages.Planning;
using PnP.Framework.Migration.Pages.Fields.Taxonomy;
using PnP.Framework.Migration.Taxonomy;
using PnP.Framework.Migration.Schema.Fields;
using System;
using System.Collections.Generic;
using System.Linq;

namespace PnP.Framework.Migration.Pages.Fields
{
    internal static class PageFieldPlanner
    {
        public static List<PageFieldAction> BuildActions(
            ClientContext targetContext,
            IEnumerable<PageFieldValueSnapshot> fields,
            ISet<string> handledFieldNames,
            ISet<string> recognizedFieldNames,
            PagePlanningOptions options,
            ICollection<TaxonomyRelationshipAction> taxonomyRelationshipActions,
            ICollection<string> blockers,
            ICollection<string> warnings,
            Microsoft.SharePoint.Client.List targetPages = null,
            bool targetPagesResolved = false,
            bool targetFieldsLoaded = false)
        {
            var pages = targetPagesResolved ? targetPages : targetContext.Web.GetPagesLibrary();
            var sourceFields = fields.OrderBy(field => field.InternalName, StringComparer.Ordinal).ToArray();
            var eligibleTaxonomyFieldNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var result = new List<PageFieldAction>();
            if (pages == null)
            {
                result.AddRange(sourceFields.Select(field => new PageFieldAction
                {
                    SourceInternalName = field.InternalName,
                    TargetInternalName = field.InternalName,
                    Disposition = PageFieldDisposition.Block,
                    Reason = "The target publishing Pages library is unavailable."
                }));
                AddTaxonomyRelationshipActions(
                    PageTaxonomyRelationshipPlanner.BuildActions(
                        targetContext,
                        pages,
                        sourceFields,
                        eligibleTaxonomyFieldNames,
                        options,
                        blockers,
                        warnings),
                    taxonomyRelationshipActions);
                return result;
            }

            if (!targetFieldsLoaded)
            {
                targetContext.Load(pages.Fields, values => values.Include(
                    field => field.InternalName,
                    field => field.TypeAsString,
                    field => field.ReadOnlyField));
                targetContext.ExecuteQueryRetry();
            }
            var targetFields = pages.Fields.ToDictionary(field => field.InternalName, StringComparer.OrdinalIgnoreCase);
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

                if (!sourceField.HasValue
                    || sourceField.Kind == PageFieldValueKind.Null
                    || (IsTaxonomy(sourceField.Kind) && sourceField.TaxonomyValues.Count == 0))
                {
                    action.Disposition = PageFieldDisposition.SkipEmpty;
                    action.Reason = "The source item has no value for this field.";
                    continue;
                }

                if (!recognizedFieldNames.Contains(sourceField.InternalName)
                    && FieldOwnershipClassifier.IsTargetRuntime(sourceField.Id, sourceField.SchemaXml))
                {
                    if (targetFields.TryGetValue(sourceField.InternalName, out var runtimeField)
                        && string.Equals(sourceField.TypeAsString, runtimeField.TypeAsString, StringComparison.OrdinalIgnoreCase))
                    {
                        action.TargetInternalName = runtimeField.InternalName;
                        action.TargetTypeAsString = runtimeField.TypeAsString;
                        action.Disposition = PageFieldDisposition.TargetRuntime;
                        action.Reason = "The source field is SharePoint-owned and an equivalent target-runtime field exists; its value is regenerated by the target runtime.";
                    }
                    else
                    {
                        action.Disposition = PageFieldDisposition.EvidenceOnly;
                        action.Reason = "The source field is SharePoint-owned, but no equivalent target-runtime field was proven; retain the captured value as evidence.";
                        warnings.Add($"SharePoint-owned field '{sourceField.InternalName}' has no proven same-type target-runtime field and remains evidence-only.");
                    }
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
                    if (IsTaxonomy(sourceField.Kind))
                    {
                        eligibleTaxonomyFieldNames.Add(sourceField.InternalName);
                        action.Disposition = PageFieldDisposition.RequiresMapping;
                        action.Reason = "The field is eligible for replay, pending an exact action for every sealed taxonomy relationship.";
                        continue;
                    }

                    if (sourceField.Required)
                    {
                        action.Disposition = PageFieldDisposition.RequiresMapping;
                        action.Reason = "The required value is captured, but its source identity must be mapped before it can be safely applied to another site.";
                        warnings.Add($"Required field '{sourceField.InternalName}' needs an explicit cross-site identity mapping.");
                    }
                    else
                    {
                        action.Disposition = PageFieldDisposition.EvidenceOnly;
                        action.Reason = "No reviewed cross-site identity mapping is available. Retain the optional source value as evidence and leave the target value unset.";
                        warnings.Add($"Optional field '{sourceField.InternalName}' has no reviewed identity mapping; its captured value remains evidence-only and the target value will be left unset.");
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

            var plannedTaxonomyRelationships = PageTaxonomyRelationshipPlanner.BuildActions(
                targetContext,
                pages,
                sourceFields,
                eligibleTaxonomyFieldNames,
                options,
                blockers,
                warnings);
            AddTaxonomyRelationshipActions(plannedTaxonomyRelationships, taxonomyRelationshipActions);
            foreach (var action in result.Where(value => eligibleTaxonomyFieldNames.Contains(value.SourceInternalName)))
            {
                var sourceField = sourceFields.Single(value =>
                    string.Equals(value.InternalName, action.SourceInternalName, StringComparison.OrdinalIgnoreCase));
                var relationships = plannedTaxonomyRelationships
                    .Where(value => value.SourceFieldId == sourceField.Id)
                    .ToArray();
                if (relationships.Length == sourceField.TaxonomyValues.Count
                    && relationships.Length > 0
                    && relationships.All(value => value.IsExecutable))
                {
                    action.Disposition = PageFieldDisposition.ApplyTaxonomyRelationships;
                    action.Reason = "Every taxonomy value has an exact reviewed relationship action. The importer will reproduce those relationships without creating or substituting Terms.";
                }
                else
                {
                    action.Disposition = PageFieldDisposition.Block;
                    action.Reason = "One or more taxonomy values lack an exact executable relationship action.";
                }
            }

            return result;
        }

        private static void AddTaxonomyRelationshipActions(
            IEnumerable<TaxonomyRelationshipAction> planned,
            ICollection<TaxonomyRelationshipAction> destination)
        {
            foreach (var action in planned)
            {
                destination.Add(action);
            }
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

        internal static bool IsTaxonomy(PageFieldValueKind kind)
        {
            return kind == PageFieldValueKind.Taxonomy || kind == PageFieldValueKind.TaxonomyCollection;
        }
    }
}
