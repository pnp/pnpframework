using Microsoft.SharePoint.Client;
using PnP.Framework.Migration.Taxonomy;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;

namespace PnP.Framework.Migration.Pages.Fields.Taxonomy
{
    internal static class PageTaxonomyRelationshipTargetInspector
    {
        public static IReadOnlyList<string> InspectHiddenListReadiness(
            ClientContext context,
            PageTaxonomyValueSnapshot sourceValue,
            TaxonomyRelationshipAction action)
        {
            if (action.Disposition == TaxonomyRelationshipDisposition.ReuseLiveInBoundTermSet)
            {
                return Array.Empty<string>();
            }

            var errors = new List<string>();
            try
            {
                var list = context.Site.RootWeb.Lists.GetByTitle("TaxonomyHiddenList");
                context.Load(list.Fields, values => values.Include(
                    value => value.InternalName,
                    value => value.TypeAsString));
                context.ExecuteQueryRetry();
                var termFields = list.Fields.Where(value => IsLanguageField(value, "Term")).Select(value => value.InternalName).ToArray();
                var pathFields = list.Fields.Where(value => IsLanguageField(value, "Path")).Select(value => value.InternalName).ToArray();
                if (termFields.Length == 0 || pathFields.Length == 0)
                {
                    return new[] { "The target TaxonomyHiddenList has no language-specific Term/Path fields." };
                }

                InspectRow(
                    context,
                    list,
                    sourceValue.Relationship.ValueHiddenListEntry,
                    action.TargetTermStoreId,
                    action.TargetValueHiddenListTermSetId.Value,
                    action.SourceTermId,
                    termFields,
                    pathFields,
                    errors);
                var catchAllSource = sourceValue.Relationship.TaxCatchAllHiddenListEntry
                    ?? sourceValue.Relationship.ValueHiddenListEntry;
                InspectRow(
                    context,
                    list,
                    catchAllSource,
                    action.TargetTermStoreId,
                    action.TargetTaxCatchAllHiddenListTermSetId.Value,
                    action.SourceTermId,
                    termFields,
                    pathFields,
                    errors);
            }
            catch (Exception exception)
            {
                errors.Add("Target TaxonomyHiddenList readiness could not be inspected: " + exception.Message);
            }
            return errors.Distinct(StringComparer.Ordinal).OrderBy(value => value, StringComparer.Ordinal).ToArray();
        }

        private static void InspectRow(
            ClientContext context,
            List list,
            TaxonomyHiddenListEntrySnapshot source,
            Guid targetStoreId,
            Guid targetSetId,
            Guid termId,
            IEnumerable<string> termFields,
            IEnumerable<string> pathFields,
            ICollection<string> errors)
        {
            foreach (var error in PageTaxonomyHiddenListLocalization.GetTargetCoverageErrors(
                         source,
                         termFields,
                         pathFields))
            {
                errors.Add(error);
            }
            if (errors.Count > 0)
            {
                return;
            }
            var rows = QueryRows(context, list, termId, targetSetId);
            if (rows.Length == 0)
            {
                return;
            }
            if (rows.Length > 1)
            {
                errors.Add($"Target TaxonomyHiddenList has multiple rows for Term '{termId:D}' and TermSet '{targetSetId:D}'.");
                return;
            }
            var targetCatchAllData = PageTaxonomySearchIdentity.Rewrite(
                source.CatchAllData,
                targetStoreId,
                targetSetId,
                termId);
            var row = rows[0];
            if (ReadGuid(row, "IdForTermStore") != targetStoreId
                || ReadGuid(row, "IdForTermSet") != targetSetId
                || ReadGuid(row, "IdForTerm") != termId
                || !string.Equals(ReadString(row, "Title"), source.Title, StringComparison.Ordinal)
                || !string.Equals(ReadString(row, "CatchAllData"), targetCatchAllData, StringComparison.Ordinal)
                || !string.Equals(ReadString(row, "CatchAllDataLabel"), source.CatchAllDataLabel, StringComparison.Ordinal)
                || !PageTaxonomyHiddenListLocalization.MatchesCapturedValues(
                    source,
                    termFields,
                    pathFields,
                    name => ReadString(row, name)))
            {
                errors.Add($"Target TaxonomyHiddenList row {row.Id} collides with the sealed relationship for Term '{termId:D}'.");
            }
        }

        private static ListItem[] QueryRows(ClientContext context, List list, Guid termId, Guid termSetId)
        {
            var query = new CamlQuery
            {
                ViewXml = "<View Scope='RecursiveAll'><Query><Where><And><Eq><FieldRef Name='IdForTerm'/><Value Type='Text'>"
                    + termId.ToString("D")
                    + "</Value></Eq><Eq><FieldRef Name='IdForTermSet'/><Value Type='Text'>"
                    + termSetId.ToString("D")
                    + "</Value></Eq></And></Where></Query><RowLimit>10</RowLimit></View>"
            };
            var items = list.GetItems(query);
            context.Load(items);
            context.ExecuteQueryRetry();
            return items.ToArray();
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
    }
}
