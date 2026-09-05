using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.Taxonomy;
using PnP.Framework.Migration.Pages.Content;
using PnP.Framework.Migration.Pages.Fields.Taxonomy;
using PnP.Framework.Migration.Taxonomy;
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
            IEnumerable<TaxonomyRelationshipAction> taxonomyRelationshipActions,
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
                    Attempted = action.WillApply,
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
                    if (action.Disposition == PageFieldDisposition.ApplyTaxonomyRelationships)
                    {
                        result.TaxonomyRelationships = PageTaxonomyRelationshipMaterializer.Ensure(
                            context,
                            field,
                            taxonomyRelationshipActions);
                        SetTaxonomyValue(
                            context,
                            targetItem,
                            action.TargetInternalName,
                            field,
                            taxonomyRelationshipActions,
                            result.TaxonomyRelationships);
                    }
                    else
                    {
                        SetValue(targetItem, action.TargetInternalName, field, replacements);
                    }
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

        private static void SetTaxonomyValue(
            ClientContext context,
            ListItem targetItem,
            string targetInternalName,
            PageFieldValueSnapshot field,
            IEnumerable<TaxonomyRelationshipAction> plannedActions,
            IEnumerable<TaxonomyRelationshipMaterializationReceipt> materializations)
        {
            var actions = plannedActions
                .Where(action => action.SourceFieldId == field.Id)
                .ToDictionary(action => RelationshipKey(action.SourceTermId, action.SourceWssId), StringComparer.Ordinal);
            var receipts = materializations
                .ToDictionary(receipt => RelationshipKey(receipt.SourceTermId, receipt.SourceWssId), receipt => receipt, StringComparer.Ordinal);
            var targetField = context.CastTo<TaxonomyField>(
                targetItem.ParentList.Fields.GetByInternalNameOrTitle(targetInternalName));
            context.Load(targetField,
                value => value.Id,
                value => value.InternalName,
                value => value.SspId,
                value => value.TermSetId,
                value => value.TextField,
                value => value.Open);
            context.ExecuteQueryRetry();
            var firstAction = actions.Values.First();
            if (targetField.Id != firstAction.TargetFieldId
                || targetField.SspId != firstAction.TargetTermStoreId
                || targetField.TermSetId != firstAction.TargetBoundTermSetId
                || targetField.TextField != firstAction.TargetTextFieldId
                || !firstAction.TargetFieldOpen.HasValue
                || targetField.Open != firstAction.TargetFieldOpen.Value)
            {
                throw new InvalidOperationException("The target taxonomy field binding changed after approval.");
            }

            var values = field.TaxonomyValues.Select(source =>
            {
                var termId = Guid.Parse(source.TermGuid);
                var action = actions[RelationshipKey(termId, source.WssId)];
                var targetWssId = action.Disposition == TaxonomyRelationshipDisposition.ReuseLiveInBoundTermSet
                    ? -1
                    : receipts[RelationshipKey(termId, source.WssId)].TargetValueWssId;
                return new TaxonomyFieldValue
                {
                    Label = source.Label,
                    TermGuid = termId.ToString("D"),
                    WssId = targetWssId
                };
            }).ToArray();
            if (field.Kind == PageFieldValueKind.Taxonomy)
            {
                targetField.SetFieldValueByValue(targetItem, values.Single());
                return;
            }

            var serialized = string.Join(";#", values.Select(value =>
                value.WssId.ToString(CultureInfo.InvariantCulture)
                + ";#"
                + value.Label
                + "|"
                + value.TermGuid));
            targetField.SetFieldValueByValueCollection(
                targetItem,
                new TaxonomyFieldValueCollection(context, serialized, targetField));
        }

        private static string RelationshipKey(Guid termId, int sourceWssId)
        {
            return termId.ToString("D") + "/" + sourceWssId;
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
