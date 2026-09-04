using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.Taxonomy;
using PnP.Framework.Migration.Taxonomy;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;

namespace PnP.Framework.Migration.Pages.Fields.Taxonomy
{
    internal static class PageTaxonomyRelationshipSnapshotReader
    {
        public static void Enrich(
            ClientContext context,
            ListItem item,
            IEnumerable<PageFieldValueSnapshot> fields,
            ICollection<string> warnings)
        {
            var taxonomyFields = (fields ?? Array.Empty<PageFieldValueSnapshot>())
                .Where(IsTaxonomyField)
                .OrderBy(field => field.InternalName, StringComparer.Ordinal)
                .ToArray();
            if (taxonomyFields.Length == 0)
            {
                return;
            }

            var capturedAt = DateTimeOffset.UtcNow;
            var taxCatchAllIds = ReadLookupIds(item, "TaxCatchAll");
            var allWssIds = taxonomyFields
                .SelectMany(field => field.TaxonomyValues ?? new List<PageTaxonomyValueSnapshot>())
                .Where(value => value != null && value.WssId > 0)
                .Select(value => value.WssId)
                .Concat(taxCatchAllIds)
                .Distinct()
                .OrderBy(value => value)
                .ToArray();
            IDictionary<int, TaxonomyHiddenListEntrySnapshot> hiddenEntries;
            try
            {
                hiddenEntries = ReadHiddenListEntries(context, allWssIds);
            }
            catch (Exception exception)
            {
                hiddenEntries = new Dictionary<int, TaxonomyHiddenListEntrySnapshot>();
                warnings.Add("TaxonomyHiddenList evidence could not be read: " + exception.Message);
            }

            foreach (var field in taxonomyFields)
            {
                try
                {
                    ReadField(context, item, field, taxCatchAllIds, hiddenEntries, capturedAt);
                }
                catch (Exception exception)
                {
                    MarkConflict(field, capturedAt, taxCatchAllIds, hiddenEntries, exception.Message);
                }

                PageTaxonomyRelationshipProof.Seal(field);
                foreach (var value in field.TaxonomyValues.Where(value => value?.Relationship?.State == TaxonomyRelationshipState.Conflict))
                {
                    var message = $"Taxonomy relationship '{field.InternalName}:{value.TermGuid}' could not be proven exactly and is non-executable.";
                    warnings.Add(message + " " + string.Join(" ", value.Relationship.Diagnostics));
                }
            }
        }

        private static void ReadField(
            ClientContext context,
            ListItem item,
            PageFieldValueSnapshot field,
            IReadOnlyCollection<int> taxCatchAllIds,
            IDictionary<int, TaxonomyHiddenListEntrySnapshot> hiddenEntries,
            DateTimeOffset capturedAt)
        {
            var sourceField = item.ParentList.Fields.GetById(field.Id);
            var taxonomyField = context.CastTo<TaxonomyField>(sourceField);
            context.Load(taxonomyField,
                value => value.Id,
                value => value.InternalName,
                value => value.SspId,
                value => value.TermSetId,
                value => value.AnchorId,
                value => value.TextField,
                value => value.Open);
            context.ExecuteQueryRetry();
            field.TaxonomyBinding = new TaxonomyFieldRelationshipBindingSnapshot
            {
                FieldId = taxonomyField.Id,
                FieldInternalName = taxonomyField.InternalName,
                TermStoreId = taxonomyField.SspId,
                BoundTermSetId = taxonomyField.TermSetId,
                AnchorTermId = taxonomyField.AnchorId,
                TextFieldId = taxonomyField.TextField,
                Open = taxonomyField.Open
            };

            var session = TaxonomySession.GetTaxonomySession(context);
            var store = session.TermStores.GetById(taxonomyField.SspId);
            foreach (var value in field.TaxonomyValues.Where(value => value != null))
            {
                Guid termId;
                hiddenEntries.TryGetValue(value.WssId, out var valueEntry);
                var relationship = new TaxonomyValueRelationshipSnapshot
                {
                    CapturedAtUtc = capturedAt,
                    State = TaxonomyRelationshipState.Unknown,
                    ValueHiddenListEntry = valueEntry
                };
                value.Relationship = relationship;
                if (!Guid.TryParse(value.TermGuid, out termId) || termId == Guid.Empty)
                {
                    relationship.State = TaxonomyRelationshipState.Conflict;
                    relationship.Diagnostics.Add("The captured taxonomy value has no valid Term GUID.");
                    continue;
                }

                var catchAllMatches = taxCatchAllIds
                    .Where(hiddenEntries.ContainsKey)
                    .Select(id => hiddenEntries[id])
                    .Where(entry => entry.TermId == termId && entry.TermSetId == taxonomyField.TermSetId)
                    .ToArray();
                if (catchAllMatches.Length == 1)
                {
                    relationship.TaxCatchAllHiddenListEntry = catchAllMatches[0];
                }
                else if (catchAllMatches.Length > 1)
                {
                    relationship.Diagnostics.Add("TaxCatchAll contains multiple bound-TermSet rows for the captured Term GUID.");
                }

                var boundTerm = store.GetTermInTermSet(taxonomyField.TermSetId, termId);
                var globalTerm = store.GetTerm(termId);
                context.Load(boundTerm,
                    term => term.Id,
                    term => term.Name,
                    term => term.PathOfTerm,
                    term => term.IsAvailableForTagging);
                context.Load(globalTerm,
                    term => term.Id,
                    term => term.Name,
                    term => term.PathOfTerm,
                    term => term.IsAvailableForTagging);
                context.ExecuteQueryRetry();
                var boundExists = !boundTerm.ServerObjectIsNull.GetValueOrDefault(true);
                var globalExists = !globalTerm.ServerObjectIsNull.GetValueOrDefault(true);
                if (boundExists)
                {
                    relationship.State = TaxonomyRelationshipState.LiveInBoundTermSet;
                    relationship.LiveTermSetId = taxonomyField.TermSetId;
                    relationship.LiveTermLabel = boundTerm.Name;
                    relationship.LiveTermPath = boundTerm.PathOfTerm;
                    relationship.LiveTermAvailableForTagging = boundTerm.IsAvailableForTagging;
                }
                else if (globalExists)
                {
                    context.Load(globalTerm.TermSet, set => set.Id, set => set.Name);
                    context.ExecuteQueryRetry();
                    relationship.State = globalTerm.TermSet.Id == taxonomyField.TermSetId
                        ? TaxonomyRelationshipState.Conflict
                        : TaxonomyRelationshipState.LiveOutsideBoundTermSet;
                    relationship.LiveTermSetId = globalTerm.TermSet.Id;
                    relationship.LiveTermSetName = globalTerm.TermSet.Name;
                    relationship.LiveTermLabel = globalTerm.Name;
                    relationship.LiveTermPath = globalTerm.PathOfTerm;
                    relationship.LiveTermAvailableForTagging = globalTerm.IsAvailableForTagging;
                    if (relationship.State == TaxonomyRelationshipState.Conflict)
                    {
                        relationship.Diagnostics.Add("The global Term reports the bound TermSet even though GetTermInTermSet did not resolve it.");
                    }
                }
                else
                {
                    relationship.State = TaxonomyRelationshipState.DanglingTermAbsent;
                }

                var fidelityErrors = PageTaxonomyRelationshipEvidence.GetFidelityErrors(field, value);
                if (fidelityErrors.Count > 0)
                {
                    relationship.Diagnostics = relationship.Diagnostics
                        .Concat(fidelityErrors)
                        .Distinct(StringComparer.Ordinal)
                        .OrderBy(message => message, StringComparer.Ordinal)
                        .ToList();
                    relationship.State = TaxonomyRelationshipState.Conflict;
                }
            }
        }

        private static IDictionary<int, TaxonomyHiddenListEntrySnapshot> ReadHiddenListEntries(
            ClientContext context,
            IReadOnlyCollection<int> wssIds)
        {
            if (wssIds.Count == 0)
            {
                return new Dictionary<int, TaxonomyHiddenListEntrySnapshot>();
            }

            var hiddenList = context.Site.RootWeb.Lists.GetByTitle("TaxonomyHiddenList");
            context.Load(hiddenList.Fields, values => values.Include(
                value => value.InternalName,
                value => value.TypeAsString));
            context.ExecuteQueryRetry();
            var termFields = hiddenList.Fields
                .Where(field => IsLanguageField(field, "Term"))
                .Select(field => field.InternalName)
                .OrderBy(value => value, StringComparer.Ordinal)
                .ToArray();
            var pathFields = hiddenList.Fields
                .Where(field => IsLanguageField(field, "Path"))
                .Select(field => field.InternalName)
                .OrderBy(value => value, StringComparer.Ordinal)
                .ToArray();
            var result = new Dictionary<int, TaxonomyHiddenListEntrySnapshot>();
            const int batchSize = 200;
            var orderedIds = wssIds.OrderBy(value => value).ToArray();
            for (var offset = 0; offset < orderedIds.Length; offset += batchSize)
            {
                var batch = orderedIds.Skip(offset).Take(batchSize).ToArray();
                var query = new CamlQuery
                {
                    ViewXml = "<View Scope='RecursiveAll'><Query><Where><In><FieldRef Name='ID'/><Values>"
                        + string.Join(string.Empty, batch.Select(id => "<Value Type='Integer'>" + id.ToString(CultureInfo.InvariantCulture) + "</Value>"))
                        + "</Values></In></Where></Query><RowLimit>200</RowLimit></View>"
                };
                var items = hiddenList.GetItems(query);
                context.Load(items);
                context.ExecuteQueryRetry();
                foreach (var value in items)
                {
                    result.Add(value.Id, new TaxonomyHiddenListEntrySnapshot
                    {
                        WssId = value.Id,
                        TermStoreId = ReadGuid(value, "IdForTermStore"),
                        TermSetId = ReadGuid(value, "IdForTermSet"),
                        TermId = ReadGuid(value, "IdForTerm"),
                        Title = ReadString(value, "Title"),
                        CatchAllData = ReadString(value, "CatchAllData"),
                        CatchAllDataLabel = ReadString(value, "CatchAllDataLabel"),
                        Terms = termFields.Select(name => new TaxonomyLocalizedTextSnapshot
                        {
                            FieldInternalName = name,
                            Value = ReadString(value, name)
                        }).ToList(),
                        Paths = pathFields.Select(name => new TaxonomyLocalizedTextSnapshot
                        {
                            FieldInternalName = name,
                            Value = ReadString(value, name)
                        }).ToList()
                    });
                }
            }
            return result;
        }

        private static void MarkConflict(
            PageFieldValueSnapshot field,
            DateTimeOffset capturedAt,
            IReadOnlyCollection<int> taxCatchAllIds,
            IDictionary<int, TaxonomyHiddenListEntrySnapshot> hiddenEntries,
            string diagnostic)
        {
            if (field.TaxonomyBinding == null)
            {
                field.TaxonomyBinding = new TaxonomyFieldRelationshipBindingSnapshot
                {
                    FieldId = field.Id,
                    FieldInternalName = field.InternalName
                };
            }
            foreach (var value in field.TaxonomyValues.Where(value => value != null))
            {
                hiddenEntries.TryGetValue(value.WssId, out var entry);
                if (value.Relationship != null
                    && value.Relationship.State != TaxonomyRelationshipState.Unknown
                    && value.Relationship.State != TaxonomyRelationshipState.Conflict
                    && PageTaxonomyRelationshipEvidence.GetFidelityErrors(field, value).Count == 0)
                {
                    continue;
                }

                var relationship = value.Relationship ?? new TaxonomyValueRelationshipSnapshot
                {
                    CapturedAtUtc = capturedAt,
                    ValueHiddenListEntry = entry
                };
                relationship.State = TaxonomyRelationshipState.Conflict;
                relationship.Diagnostics = (relationship.Diagnostics ?? new List<string>())
                    .Concat(new[] { diagnostic })
                    .Where(valueDiagnostic => !string.IsNullOrWhiteSpace(valueDiagnostic))
                    .Distinct(StringComparer.Ordinal)
                    .OrderBy(valueDiagnostic => valueDiagnostic, StringComparer.Ordinal)
                    .ToList();
                Guid termId;
                if (relationship.TaxCatchAllHiddenListEntry == null
                    && Guid.TryParse(value.TermGuid, out termId))
                {
                    var matches = taxCatchAllIds
                        .Where(hiddenEntries.ContainsKey)
                        .Select(id => hiddenEntries[id])
                        .Where(candidate => candidate.TermId == termId
                            && (field.TaxonomyBinding.BoundTermSetId == Guid.Empty
                                || candidate.TermSetId == field.TaxonomyBinding.BoundTermSetId))
                        .ToArray();
                    if (matches.Length == 1)
                    {
                        relationship.TaxCatchAllHiddenListEntry = matches[0];
                    }
                    else if (matches.Length > 1)
                    {
                        relationship.Diagnostics.Add("TaxCatchAll contains multiple candidate rows for the captured Term GUID.");
                    }
                }
                value.Relationship = relationship;
            }
        }

        private static bool IsTaxonomyField(PageFieldValueSnapshot field)
        {
            return field != null
                && (field.Kind == PageFieldValueKind.Taxonomy
                    || field.Kind == PageFieldValueKind.TaxonomyCollection);
        }

        private static IReadOnlyCollection<int> ReadLookupIds(ListItem item, string internalName)
        {
            if (!item.FieldValues.TryGetValue(internalName, out var value) || value == null)
            {
                return Array.Empty<int>();
            }
            if (value is FieldLookupValue single)
            {
                return new[] { single.LookupId };
            }
            if (value is FieldLookupValue[] values)
            {
                return values.Select(itemValue => itemValue.LookupId).Where(id => id > 0).Distinct().ToArray();
            }
            return Array.Empty<int>();
        }

        private static bool IsLanguageField(Field field, string prefix)
        {
            if (field == null
                || !string.Equals(field.TypeAsString, "Text", StringComparison.OrdinalIgnoreCase)
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
