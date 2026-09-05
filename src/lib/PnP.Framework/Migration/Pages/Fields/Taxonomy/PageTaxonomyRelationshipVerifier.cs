using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.Taxonomy;
using PnP.Framework.Migration.Taxonomy;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;

namespace PnP.Framework.Migration.Pages.Fields.Taxonomy
{
    internal static class PageTaxonomyRelationshipVerifier
    {
        public static IList<TaxonomyRelationshipVerificationResult> Verify(
            ClientContext context,
            List targetList,
            ListItem item,
            IEnumerable<PageFieldValueSnapshot> fields,
            IEnumerable<TaxonomyRelationshipAction> actions,
            IEnumerable<PageFieldImportResult> fieldResults)
        {
            var fieldById = fields.ToDictionary(field => field.Id);
            var receipts = fieldResults
                .SelectMany(result => result.TaxonomyRelationships ?? new List<TaxonomyRelationshipMaterializationReceipt>())
                .ToDictionary(
                    receipt => Key(receipt.SourceFieldId, receipt.SourceTermId, receipt.SourceWssId),
                    StringComparer.Ordinal);
            var taxCatchAllIds = ReadLookupIds(item, "TaxCatchAll");
            var results = new List<TaxonomyRelationshipVerificationResult>();
            foreach (var action in actions
                         .Where(value => value.IsExecutable)
                         .OrderBy(value => value.SourceFieldInternalName, StringComparer.Ordinal)
                         .ThenBy(value => value.SourceTermId))
            {
                var result = new TaxonomyRelationshipVerificationResult
                {
                    SourceFieldId = action.SourceFieldId,
                    SourceFieldInternalName = action.SourceFieldInternalName,
                    SourceTermId = action.SourceTermId,
                    Disposition = action.Disposition
                };
                results.Add(result);
                try
                {
                    var field = fieldById[action.SourceFieldId];
                    var sourceValue = field.TaxonomyValues.Single(value =>
                        ParseTermId(value.TermGuid) == action.SourceTermId
                        && value.WssId == action.SourceWssId);
                    var observed = ReadTaxonomyValues(item, action.SourceFieldInternalName)
                        .Where(value => ParseTermId(value.TermGuid) == action.SourceTermId)
                        .ToArray();
                    result.PageValueMatched = observed.Length == 1
                        && string.Equals(observed[0].Label, sourceValue.Label, StringComparison.Ordinal);
                    result.ObservedWssId = observed.Length == 1 ? observed[0].WssId : 0;
                    result.RelationshipStateMatched = VerifyRelationshipState(context, targetList, field, sourceValue, action);
                    if (action.Disposition == TaxonomyRelationshipDisposition.ReuseLiveInBoundTermSet)
                    {
                        result.HiddenListIdentityMatched = observed.Length == 1
                            && VerifyLiveBoundHiddenRow(context, observed[0].WssId, action);
                        result.TaxCatchAllMatched = observed.Length == 1
                            && taxCatchAllIds.Contains(observed[0].WssId);
                    }
                    else if (observed.Length == 1
                        && receipts.TryGetValue(Key(action.SourceFieldId, action.SourceTermId, action.SourceWssId), out var receipt))
                    {
                        result.HiddenListIdentityMatched = VerifyObservedHiddenRow(
                            context,
                            observed[0].WssId,
                            field,
                            sourceValue,
                            action,
                            receipt);
                        result.TaxCatchAllMatched = taxCatchAllIds.Contains(receipt.TargetTaxCatchAllWssId);
                    }
                    result.Message = result.Passed
                        ? "Fresh readback reproduced the sealed taxonomy relationship."
                        : "Fresh readback differs from the sealed taxonomy relationship.";
                }
                catch (Exception exception)
                {
                    result.Message = exception.Message;
                }
            }
            return results;
        }

        private static bool VerifyRelationshipState(
            ClientContext context,
            List targetList,
            PageFieldValueSnapshot sourceField,
            PageTaxonomyValueSnapshot sourceValue,
            TaxonomyRelationshipAction action)
        {
            var targetField = context.CastTo<TaxonomyField>(targetList.Fields.GetByInternalNameOrTitle(action.SourceFieldInternalName));
            context.Load(targetField,
                field => field.Id,
                field => field.SspId,
                field => field.TermSetId,
                field => field.TextField,
                field => field.Open);
            var session = TaxonomySession.GetTaxonomySession(context);
            var store = session.TermStores.GetById(action.TargetTermStoreId);
            var inBound = store.GetTermInTermSet(action.TargetBoundTermSetId, action.SourceTermId);
            var global = store.GetTerm(action.SourceTermId);
            context.Load(inBound,
                term => term.Id,
                term => term.Name,
                term => term.PathOfTerm,
                term => term.IsAvailableForTagging);
            context.Load(global,
                term => term.Id,
                term => term.Name,
                term => term.PathOfTerm,
                term => term.IsAvailableForTagging);
            context.ExecuteQueryRetry();
            if (targetField.Id != action.TargetFieldId
                || targetField.SspId != action.TargetTermStoreId
                || targetField.TermSetId != action.TargetBoundTermSetId
                || targetField.TextField != action.TargetTextFieldId
                || !action.TargetFieldOpen.HasValue
                || targetField.Open != action.TargetFieldOpen.Value
                || targetField.Open != sourceField.TaxonomyBinding.Open)
            {
                return false;
            }
            var inBoundExists = !inBound.ServerObjectIsNull.GetValueOrDefault(true);
            var globalExists = !global.ServerObjectIsNull.GetValueOrDefault(true);
            switch (action.Disposition)
            {
                case TaxonomyRelationshipDisposition.ReuseLiveInBoundTermSet:
                    return inBoundExists
                        && string.Equals(inBound.Name, sourceValue.Relationship.LiveTermLabel, StringComparison.Ordinal)
                        && string.Equals(inBound.PathOfTerm, sourceValue.Relationship.LiveTermPath, StringComparison.Ordinal)
                        && inBound.IsAvailableForTagging == sourceValue.Relationship.LiveTermAvailableForTagging.Value;
                case TaxonomyRelationshipDisposition.PreserveDanglingTermAbsent:
                    return !inBoundExists && !globalExists;
                case TaxonomyRelationshipDisposition.PreserveLiveOutsideBoundTermSet:
                    if (inBoundExists || !globalExists)
                    {
                        return false;
                    }
                    context.Load(global.TermSet, set => set.Id, set => set.Name);
                    context.ExecuteQueryRetry();
                    return global.TermSet.Id == action.TargetLiveTermSetId
                        && string.Equals(global.Name, sourceValue.Relationship.LiveTermLabel, StringComparison.Ordinal)
                        && string.Equals(global.PathOfTerm, sourceValue.Relationship.LiveTermPath, StringComparison.Ordinal)
                        && global.IsAvailableForTagging == sourceValue.Relationship.LiveTermAvailableForTagging.Value;
                default:
                    return false;
            }
        }

        private static bool VerifyLiveBoundHiddenRow(
            ClientContext context,
            int observedWssId,
            TaxonomyRelationshipAction action)
        {
            if (observedWssId <= 0)
            {
                return false;
            }
            var list = context.Site.RootWeb.Lists.GetByTitle("TaxonomyHiddenList");
            var row = list.GetItemById(observedWssId);
            context.Load(row);
            context.ExecuteQueryRetry();
            return row.Id == observedWssId
                && ReadGuid(row, "IdForTermStore") == action.TargetTermStoreId
                && ReadGuid(row, "IdForTermSet") == action.TargetBoundTermSetId
                && ReadGuid(row, "IdForTerm") == action.SourceTermId;
        }

        private static bool VerifyObservedHiddenRow(
            ClientContext context,
            int observedWssId,
            PageFieldValueSnapshot field,
            PageTaxonomyValueSnapshot sourceValue,
            TaxonomyRelationshipAction action,
            TaxonomyRelationshipMaterializationReceipt receipt)
        {
            TaxonomyHiddenListEntrySnapshot sourceEntry;
            Guid expectedSetId;
            if (observedWssId == receipt.TargetValueWssId)
            {
                sourceEntry = sourceValue.Relationship.ValueHiddenListEntry;
                expectedSetId = action.TargetValueHiddenListTermSetId.Value;
            }
            else if (observedWssId == receipt.TargetTaxCatchAllWssId)
            {
                sourceEntry = sourceValue.Relationship.TaxCatchAllHiddenListEntry
                    ?? sourceValue.Relationship.ValueHiddenListEntry;
                expectedSetId = action.TargetTaxCatchAllHiddenListTermSetId.Value;
            }
            else
            {
                return false;
            }

            var list = context.Site.RootWeb.Lists.GetByTitle("TaxonomyHiddenList");
            context.Load(list.Fields, values => values.Include(
                value => value.InternalName,
                value => value.TypeAsString));
            var row = list.GetItemById(observedWssId);
            context.Load(row);
            context.ExecuteQueryRetry();
            var termFields = list.Fields.Where(value => IsLanguageField(value, "Term")).Select(value => value.InternalName).ToArray();
            var pathFields = list.Fields.Where(value => IsLanguageField(value, "Path")).Select(value => value.InternalName).ToArray();
            var expectedCatchAllData = PageTaxonomySearchIdentity.Rewrite(
                sourceEntry.CatchAllData,
                action.TargetTermStoreId,
                expectedSetId,
                action.SourceTermId);
            return row.Id == observedWssId
                && ReadGuid(row, "IdForTermStore") == action.TargetTermStoreId
                && ReadGuid(row, "IdForTermSet") == expectedSetId
                && ReadGuid(row, "IdForTerm") == action.SourceTermId
                && string.Equals(ReadString(row, "Title"), sourceEntry.Title, StringComparison.Ordinal)
                && string.Equals(ReadString(row, "CatchAllData"), expectedCatchAllData, StringComparison.Ordinal)
                && string.Equals(ReadString(row, "CatchAllDataLabel"), sourceEntry.CatchAllDataLabel, StringComparison.Ordinal)
                && PageTaxonomyHiddenListLocalization.GetTargetCoverageErrors(sourceEntry, termFields, pathFields).Count == 0
                && PageTaxonomyHiddenListLocalization.MatchesCapturedValues(
                    sourceEntry,
                    termFields,
                    pathFields,
                    name => ReadString(row, name));
        }

        private static IEnumerable<TaxonomyFieldValue> ReadTaxonomyValues(ListItem item, string fieldName)
        {
            if (!item.FieldValues.TryGetValue(fieldName, out var value) || value == null)
            {
                return Array.Empty<TaxonomyFieldValue>();
            }
            if (value is TaxonomyFieldValue single)
            {
                return new[] { single };
            }
            if (value is TaxonomyFieldValueCollection collection)
            {
                return collection.ToArray();
            }
            return Array.Empty<TaxonomyFieldValue>();
        }

        private static ISet<int> ReadLookupIds(ListItem item, string fieldName)
        {
            if (!item.FieldValues.TryGetValue(fieldName, out var value) || value == null)
            {
                return new HashSet<int>();
            }
            if (value is FieldLookupValue single)
            {
                return new HashSet<int> { single.LookupId };
            }
            if (value is FieldLookupValue[] values)
            {
                return new HashSet<int>(values.Select(itemValue => itemValue.LookupId));
            }
            return new HashSet<int>();
        }

        private static bool IsLanguageField(Field field, string prefix)
        {
            if (!string.Equals(field.TypeAsString, "Text", StringComparison.OrdinalIgnoreCase)
                || !field.InternalName.StartsWith(prefix, StringComparison.Ordinal)
                || field.InternalName.Length == prefix.Length)
            {
                return false;
            }
            int ignored;
            return int.TryParse(field.InternalName.Substring(prefix.Length), NumberStyles.Integer, CultureInfo.InvariantCulture, out ignored);
        }

        private static Guid ReadGuid(ListItem item, string fieldName)
        {
            Guid result;
            return Guid.TryParse(ReadString(item, fieldName), out result) ? result : Guid.Empty;
        }

        private static string ReadString(ListItem item, string fieldName)
        {
            return item.FieldValues.TryGetValue(fieldName, out var value)
                ? Convert.ToString(value, CultureInfo.InvariantCulture) ?? string.Empty
                : string.Empty;
        }

        private static Guid ParseTermId(string value)
        {
            Guid result;
            return Guid.TryParse(value, out result) ? result : Guid.Empty;
        }

        private static string Key(Guid fieldId, Guid termId, int sourceWssId)
        {
            return fieldId.ToString("D") + "/" + termId.ToString("D") + "/" + sourceWssId;
        }
    }
}
