using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.Taxonomy;
using PnP.Framework.Migration.Taxonomy;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;

namespace PnP.Framework.Migration.Pages.Fields.Taxonomy
{
    internal static class PageTaxonomyRelationshipMaterializer
    {
        public static IList<TaxonomyRelationshipMaterializationReceipt> Ensure(
            ClientContext context,
            PageFieldValueSnapshot field,
            IEnumerable<TaxonomyRelationshipAction> plannedActions)
        {
            var actions = (plannedActions ?? Array.Empty<TaxonomyRelationshipAction>())
                .Where(action => action.SourceFieldId == field.Id)
                .OrderBy(action => action.SourceTermId)
                .ThenBy(action => action.SourceWssId)
                .ToArray();
            if (actions.Length != field.TaxonomyValues.Count || actions.Any(action => !action.IsExecutable))
            {
                throw new InvalidOperationException($"Field '{field.InternalName}' has no complete executable taxonomy relationship plan.");
            }
            var pages = context.Web.GetPagesLibrary();
            if (pages == null)
            {
                throw new InvalidOperationException("The target publishing Pages library is unavailable.");
            }
            var targetField = context.CastTo<TaxonomyField>(
                pages.Fields.GetByInternalNameOrTitle(field.InternalName));
            context.Load(targetField,
                value => value.Id,
                value => value.SspId,
                value => value.TermSetId,
                value => value.TextField,
                value => value.Open);
            context.ExecuteQueryRetry();
            if (actions.Any(action => targetField.Id != action.TargetFieldId
                || targetField.SspId != action.TargetTermStoreId
                || targetField.TermSetId != action.TargetBoundTermSetId
                || targetField.TextField != action.TargetTextFieldId
                || !action.TargetFieldOpen.HasValue
                || targetField.Open != action.TargetFieldOpen.Value))
            {
                throw new InvalidOperationException($"Target taxonomy field '{field.InternalName}' changed binding after approval.");
            }

            var values = field.TaxonomyValues.ToDictionary(
                value => Key(ParseTermId(value.TermGuid), value.WssId),
                StringComparer.Ordinal);
            var invalidActions = actions.Where(action =>
                    action.Disposition == TaxonomyRelationshipDisposition.PreserveDanglingTermAbsent
                    || action.Disposition == TaxonomyRelationshipDisposition.PreserveLiveOutsideBoundTermSet)
                .ToArray();
            HiddenListContext hidden = null;
            if (invalidActions.Length > 0)
            {
                hidden = LoadHiddenList(context);
            }

            var receipts = new List<TaxonomyRelationshipMaterializationReceipt>();
            foreach (var action in actions)
            {
                if (!values.TryGetValue(Key(action.SourceTermId, action.SourceWssId), out var value)
                    || value.Relationship == null
                    || !string.Equals(value.Relationship.EvidenceSha256, action.SourceEvidenceSha256, StringComparison.OrdinalIgnoreCase))
                {
                    throw new InvalidOperationException($"Taxonomy action '{field.InternalName}:{action.SourceTermId:D}' no longer matches one exact source relationship proof.");
                }

                AssertTargetRelationshipState(context, value, action);
                if (action.Disposition == TaxonomyRelationshipDisposition.ReuseLiveInBoundTermSet)
                {
                    receipts.Add(new TaxonomyRelationshipMaterializationReceipt
                    {
                        SourceFieldId = field.Id,
                        SourceTermId = action.SourceTermId,
                        SourceWssId = action.SourceWssId,
                        Disposition = action.Disposition,
                        TargetValueWssId = -1,
                        TargetTaxCatchAllWssId = -1,
                        TargetRelationshipStateVerified = true,
                        HiddenListIdentityVerified = true,
                        Message = "Reused the exact live Term in the mapped bound TermSet; no Term was created or substituted."
                    });
                    continue;
                }

                var sourceValueEntry = value.Relationship.ValueHiddenListEntry;
                var sourceCatchAllEntry = value.Relationship.TaxCatchAllHiddenListEntry ?? sourceValueEntry;
                var valueRow = EnsureRow(
                    context,
                    hidden,
                    action.TargetTermStoreId,
                    action.TargetValueHiddenListTermSetId.Value,
                    action.SourceTermId,
                    sourceValueEntry);
                var catchAllRow = action.TargetTaxCatchAllHiddenListTermSetId == action.TargetValueHiddenListTermSetId
                    ? valueRow
                    : EnsureRow(
                        context,
                        hidden,
                        action.TargetTermStoreId,
                        action.TargetTaxCatchAllHiddenListTermSetId.Value,
                        action.SourceTermId,
                        sourceCatchAllEntry);
                receipts.Add(new TaxonomyRelationshipMaterializationReceipt
                {
                    SourceFieldId = field.Id,
                    SourceTermId = action.SourceTermId,
                    SourceWssId = action.SourceWssId,
                    Disposition = action.Disposition,
                    TargetValueWssId = valueRow.WssId,
                    TargetTaxCatchAllWssId = catchAllRow.WssId,
                    ChangedTarget = valueRow.Created || catchAllRow.Created,
                    TargetRelationshipStateVerified = true,
                    HiddenListIdentityVerified = true,
                    Message = action.Disposition == TaxonomyRelationshipDisposition.PreserveDanglingTermAbsent
                        ? "Created or reused an exact target-local dangling taxonomy relationship; the Term remains absent."
                        : "Created or reused the exact target-local live-outside-bound relationship and UNVALIDATED TaxCatchAll identity."
                });
            }

            return receipts;
        }

        private static void AssertTargetRelationshipState(
            ClientContext context,
            PageTaxonomyValueSnapshot value,
            TaxonomyRelationshipAction action)
        {
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
            var inBoundExists = !inBound.ServerObjectIsNull.GetValueOrDefault(true);
            var globalExists = !global.ServerObjectIsNull.GetValueOrDefault(true);
            switch (action.Disposition)
            {
                case TaxonomyRelationshipDisposition.ReuseLiveInBoundTermSet:
                    if (!inBoundExists
                        || !string.Equals(inBound.Name, value.Relationship.LiveTermLabel, StringComparison.Ordinal)
                        || !string.Equals(inBound.PathOfTerm, value.Relationship.LiveTermPath, StringComparison.Ordinal)
                        || inBound.IsAvailableForTagging != value.Relationship.LiveTermAvailableForTagging.Value)
                    {
                        throw new InvalidOperationException("The target live-in-bound taxonomy relationship changed after approval.");
                    }
                    break;
                case TaxonomyRelationshipDisposition.PreserveDanglingTermAbsent:
                    if (inBoundExists || globalExists)
                    {
                        throw new InvalidOperationException("The target now resolves the source dangling Term GUID live; preserving it would heal the relationship.");
                    }
                    break;
                case TaxonomyRelationshipDisposition.PreserveLiveOutsideBoundTermSet:
                    if (inBoundExists || !globalExists)
                    {
                        throw new InvalidOperationException("The target live-outside-bound taxonomy relationship changed after approval.");
                    }
                    context.Load(global.TermSet, set => set.Id, set => set.Name);
                    context.ExecuteQueryRetry();
                    if (global.TermSet.Id != action.TargetLiveTermSetId
                        || !string.Equals(global.Name, value.Relationship.LiveTermLabel, StringComparison.Ordinal)
                        || !string.Equals(global.PathOfTerm, value.Relationship.LiveTermPath, StringComparison.Ordinal)
                        || global.IsAvailableForTagging != value.Relationship.LiveTermAvailableForTagging.Value)
                    {
                        throw new InvalidOperationException("The target global Term no longer exactly matches the sealed live-outside-bound source relationship.");
                    }
                    break;
                default:
                    throw new InvalidOperationException($"Taxonomy relationship disposition '{action.Disposition}' is not executable.");
            }
        }

        private static HiddenListContext LoadHiddenList(ClientContext context)
        {
            var list = context.Site.RootWeb.Lists.GetByTitle("TaxonomyHiddenList");
            context.Load(list, value => value.Id);
            context.Load(list.Fields, values => values.Include(
                value => value.InternalName,
                value => value.TypeAsString));
            context.ExecuteQueryRetry();
            var termFields = list.Fields
                .Where(field => IsLanguageField(field, "Term"))
                .Select(field => field.InternalName)
                .OrderBy(value => value, StringComparer.Ordinal)
                .ToArray();
            var pathFields = list.Fields
                .Where(field => IsLanguageField(field, "Path"))
                .Select(field => field.InternalName)
                .OrderBy(value => value, StringComparer.Ordinal)
                .ToArray();
            if (termFields.Length == 0 || pathFields.Length == 0)
            {
                throw new InvalidOperationException("The target TaxonomyHiddenList has no language-specific Term/Path fields.");
            }
            return new HiddenListContext(list, termFields, pathFields);
        }

        private static MaterializedRow EnsureRow(
            ClientContext context,
            HiddenListContext hidden,
            Guid targetStoreId,
            Guid targetSetId,
            Guid termId,
            TaxonomyHiddenListEntrySnapshot source)
        {
            if (source == null)
            {
                throw new InvalidOperationException($"Term '{termId:D}' has no sealed source TaxonomyHiddenList row.");
            }
            var coverageErrors = PageTaxonomyHiddenListLocalization.GetTargetCoverageErrors(
                source,
                hidden.TermFields,
                hidden.PathFields);
            if (coverageErrors.Count > 0)
            {
                throw new InvalidOperationException(string.Join(" ", coverageErrors));
            }
            var targetCatchAllData = PageTaxonomySearchIdentity.Rewrite(
                source.CatchAllData,
                targetStoreId,
                targetSetId,
                termId);
            var matches = QueryRows(context, hidden.List, termId, targetSetId);
            if (matches.Length > 1)
            {
                throw new InvalidOperationException($"Target TaxonomyHiddenList has multiple rows for Term '{termId:D}' and TermSet '{targetSetId:D}'.");
            }

            var created = false;
            if (matches.Length == 0)
            {
                var item = hidden.List.AddItem(new ListItemCreationInformation());
                item["IdForTerm"] = termId.ToString("D");
                item["IdForTermSet"] = targetSetId.ToString("D");
                item["IdForTermStore"] = targetStoreId.ToString("D");
                item["Title"] = source.Title;
                item["CatchAllData"] = targetCatchAllData;
                item["CatchAllDataLabel"] = source.CatchAllDataLabel;
                foreach (var localized in source.Terms)
                {
                    item[localized.FieldInternalName] = localized.Value ?? string.Empty;
                }
                foreach (var localized in source.Paths)
                {
                    item[localized.FieldInternalName] = localized.Value ?? string.Empty;
                }
                item.Update();
                context.ExecuteQueryRetry();
                created = true;
                matches = QueryRows(context, hidden.List, termId, targetSetId);
            }

            if (matches.Length != 1
                || !IsExactRow(
                    matches[0],
                    targetStoreId,
                    targetSetId,
                    termId,
                    source,
                    targetCatchAllData,
                    hidden.TermFields,
                    hidden.PathFields))
            {
                throw new InvalidOperationException($"Target TaxonomyHiddenList row collides with the sealed relationship for Term '{termId:D}'.");
            }
            return new MaterializedRow(matches[0].Id, created);
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

        private static bool IsExactRow(
            ListItem item,
            Guid storeId,
            Guid setId,
            Guid termId,
            TaxonomyHiddenListEntrySnapshot source,
            string catchAllData,
            IEnumerable<string> termFields,
            IEnumerable<string> pathFields)
        {
            return item.Id > 0
                && ReadGuid(item, "IdForTermStore") == storeId
                && ReadGuid(item, "IdForTermSet") == setId
                && ReadGuid(item, "IdForTerm") == termId
                && string.Equals(ReadString(item, "Title"), source.Title, StringComparison.Ordinal)
                && string.Equals(ReadString(item, "CatchAllData"), catchAllData, StringComparison.Ordinal)
                && string.Equals(ReadString(item, "CatchAllDataLabel"), source.CatchAllDataLabel, StringComparison.Ordinal)
                && PageTaxonomyHiddenListLocalization.MatchesCapturedValues(
                    source,
                    termFields,
                    pathFields,
                    name => ReadString(item, name));
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

        private static string Key(Guid termId, int wssId)
        {
            return termId.ToString("D") + "/" + wssId;
        }

        private sealed class HiddenListContext
        {
            public HiddenListContext(List list, string[] termFields, string[] pathFields)
            {
                List = list;
                TermFields = termFields;
                PathFields = pathFields;
            }

            public List List { get; }

            public string[] TermFields { get; }

            public string[] PathFields { get; }
        }

        private sealed class MaterializedRow
        {
            public MaterializedRow(int wssId, bool created)
            {
                WssId = wssId;
                Created = created;
            }

            public int WssId { get; }

            public bool Created { get; }
        }
    }
}
