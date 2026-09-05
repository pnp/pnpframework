using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.Taxonomy;
using PnP.Framework.Migration.Lists.Items;
using PnP.Framework.Migration.Lists.Planning;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;

namespace PnP.Framework.Migration.Lists.Execution
{
    internal static class ListItemValueWriter
    {
        public static void Apply(
            ClientContext context,
            List targetList,
            ListItem targetItem,
            ListItemSnapshot sourceItem,
            ListMaterializationPlan plan,
            IDictionary<Guid, ListMaterializationReceipt> dependencyReceipts,
            IDictionary<string, string> contentTypeIds,
            bool lookupPhase)
        {
            var plans = plan.Fields.ToDictionary(value => value.InternalName, StringComparer.OrdinalIgnoreCase);
            var changed = false;
            foreach (var value in sourceItem.Values)
            {
                if (!lookupPhase && string.Equals(value.InternalName, "ContentTypeId", StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }
                ListFieldMaterializationPlan fieldPlan;
                if (!plans.TryGetValue(value.InternalName, out fieldPlan) || value.Kind == ListItemValueKind.Null)
                {
                    continue;
                }
                var isLookup = fieldPlan.Disposition == ListFieldMaterializationDisposition.MapLookup;
                if (isLookup != lookupPhase)
                {
                    continue;
                }
                if (!WritesValue(fieldPlan.Disposition))
                {
                    continue;
                }
                if (fieldPlan.Disposition == ListFieldMaterializationDisposition.MapTaxonomy)
                {
                    var field = context.CastTo<TaxonomyField>(targetList.Fields.GetById(fieldPlan.SourceFieldId));
                    field.SetFieldValueByLabelGuidPair(targetItem, string.Join(";", value.TaxonomyValues.Select(term => term.Label + "|" + term.TermGuid)));
                    continue;
                }
                targetItem[fieldPlan.InternalName] = ToTargetValue(value, fieldPlan, dependencyReceipts);
                changed = true;
            }
            if (changed)
            {
                targetItem.Update();
                context.ExecuteQueryRetry();
            }
        }

        internal static bool ApplyContentType(
            ClientContext context,
            ListItem targetItem,
            ListItemSnapshot sourceItem,
            IDictionary<string, string> contentTypeIds)
        {
            var value = sourceItem.Values.FirstOrDefault(candidate =>
                string.Equals(candidate.InternalName, "ContentTypeId", StringComparison.OrdinalIgnoreCase));
            var sourceContentTypeId = value?.ScalarValue ?? value?.RawValue;
            if (string.IsNullOrWhiteSpace(sourceContentTypeId)
                || contentTypeIds == null
                || !contentTypeIds.TryGetValue(sourceContentTypeId, out var targetContentTypeId))
            {
                return false;
            }

            // SharePoint may apply Content Type defaults after other field values
            // when both are submitted in one request. Commit the Content Type
            // first so authored values such as Title and Order are not reset to
            // their document-library defaults.
            targetItem["ContentTypeId"] = targetContentTypeId;
            targetItem.Update();
            context.ExecuteQueryRetry();
            return true;
        }

        private static bool WritesValue(ListFieldMaterializationDisposition value)
        {
            return value == ListFieldMaterializationDisposition.RequireTargetRuntimeAndCopyValue
                || value == ListFieldMaterializationDisposition.CreateOrReuseOwnedAndCopyValue
                || value == ListFieldMaterializationDisposition.MapLookup
                || value == ListFieldMaterializationDisposition.MapTaxonomy;
        }

        private static object ToTargetValue(
            ListItemValueSnapshot value,
            ListFieldMaterializationPlan fieldPlan,
            IDictionary<Guid, ListMaterializationReceipt> dependencyReceipts)
        {
            switch (value.Kind)
            {
                case ListItemValueKind.String: return value.ScalarValue;
                case ListItemValueKind.StringCollection: return value.StringValues.ToArray();
                case ListItemValueKind.Boolean: return string.Equals(value.ScalarValue, "true", StringComparison.OrdinalIgnoreCase);
                case ListItemValueKind.Number:
                    decimal number;
                    return decimal.TryParse(value.ScalarValue, NumberStyles.Any, CultureInfo.InvariantCulture, out number) ? (object)number : value.ScalarValue;
                case ListItemValueKind.DateTime:
                    DateTime date;
                    return DateTime.TryParse(value.ScalarValue, CultureInfo.InvariantCulture, DateTimeStyles.RoundtripKind, out date) ? (object)date : value.ScalarValue;
                case ListItemValueKind.Guid:
                    Guid guid;
                    return Guid.TryParse(value.ScalarValue, out guid) ? (object)guid : value.ScalarValue;
                case ListItemValueKind.Url:
                    return new FieldUrlValue { Url = value.UrlValue == null ? null : value.UrlValue.Url, Description = value.UrlValue == null ? null : value.UrlValue.Description };
                case ListItemValueKind.Lookup:
                case ListItemValueKind.LookupCollection:
                    ListMaterializationReceipt receipt;
                    if (!fieldPlan.SourceLookupListId.HasValue || !dependencyReceipts.TryGetValue(fieldPlan.SourceLookupListId.Value, out receipt))
                    {
                        throw new InvalidDataException("Lookup value has no target item-ID catalog for field '" + fieldPlan.InternalName + "'.");
                    }
                    var mapped = value.LookupValues.Select(source =>
                    {
                        int targetId;
                        if (!receipt.TargetItemIds.TryGetValue(source.LookupId, out targetId))
                        {
                            throw new InvalidDataException("Lookup source item " + source.LookupId + " has no target item mapping for field '" + fieldPlan.InternalName + "'.");
                        }
                        return new FieldLookupValue { LookupId = targetId };
                    }).ToArray();
                    return value.Kind == ListItemValueKind.Lookup ? (object)mapped.Single() : mapped;
                default:
                    throw new InvalidDataException("Value kind '" + value.Kind + "' is not approved for List item replay.");
            }
        }
    }
}
