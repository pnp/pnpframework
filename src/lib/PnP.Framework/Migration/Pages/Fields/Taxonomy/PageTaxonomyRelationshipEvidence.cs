using PnP.Framework.Migration.Taxonomy;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;

namespace PnP.Framework.Migration.Pages.Fields.Taxonomy
{
    internal static class PageTaxonomyRelationshipEvidence
    {
        public static IReadOnlyList<string> ValidateSealedField(PageFieldValueSnapshot field)
        {
            var errors = new List<string>();
            if (field == null)
            {
                errors.Add("The taxonomy field snapshot is null.");
                return errors;
            }
            if (field.TaxonomyBinding == null
                || field.TaxonomyBinding.FieldId != field.Id
                || !string.Equals(field.TaxonomyBinding.FieldInternalName, field.InternalName, StringComparison.OrdinalIgnoreCase))
            {
                errors.Add($"Field '{field.InternalName}' has no structurally valid taxonomy binding evidence.");
            }

            var expectedValueSetDigest = PageTaxonomyRelationshipProof.ComputeFieldValueSetSha256(field);
            if (!string.Equals(expectedValueSetDigest, field.TaxonomyValueSetSha256, StringComparison.OrdinalIgnoreCase))
            {
                errors.Add($"Field '{field.InternalName}' taxonomy value-set digest does not match its captured values and binding.");
            }

            foreach (var value in field.TaxonomyValues ?? new List<PageTaxonomyValueSnapshot>())
            {
                if (value == null || value.Relationship == null)
                {
                    errors.Add($"Field '{field.InternalName}' contains a taxonomy value without relationship evidence.");
                    continue;
                }
                if (!string.Equals(value.Relationship.SchemaVersion, "pnp-taxonomy-value-relationship/v1", StringComparison.Ordinal)
                    || value.Relationship.CapturedAtUtc == default(DateTimeOffset)
                    || value.Relationship.Diagnostics == null)
                {
                    errors.Add($"Field '{field.InternalName}' contains malformed taxonomy relationship evidence.");
                    continue;
                }
                if (!string.Equals(value.Relationship.SourceFieldValueSetSha256, field.TaxonomyValueSetSha256, StringComparison.OrdinalIgnoreCase)
                    || !string.Equals(value.Relationship.EvidenceSha256, PageTaxonomyRelationshipProof.ComputeEvidenceSha256(field, value), StringComparison.OrdinalIgnoreCase))
                {
                    errors.Add($"Field '{field.InternalName}' Term '{value.TermGuid}' relationship proof is not sealed to the exact field value set.");
                }
                if (value.Relationship.State == TaxonomyRelationshipState.Conflict
                    && !value.Relationship.Diagnostics.Any(item => !string.IsNullOrWhiteSpace(item)))
                {
                    errors.Add($"Field '{field.InternalName}' Term '{value.TermGuid}' is marked Conflict without diagnostics.");
                }
                if (value.Relationship.State == TaxonomyRelationshipState.Unknown
                    && !value.Relationship.Diagnostics.Any(item => !string.IsNullOrWhiteSpace(item)))
                {
                    errors.Add($"Field '{field.InternalName}' Term '{value.TermGuid}' is marked Unknown without diagnostics.");
                }
                if (value.Relationship.State != TaxonomyRelationshipState.Unknown
                    && value.Relationship.State != TaxonomyRelationshipState.Conflict)
                {
                    var fidelityErrors = GetFidelityErrors(field, value);
                    if (fidelityErrors.Count > 0)
                    {
                        errors.Add($"Field '{field.InternalName}' Term '{value.TermGuid}' claims relationship state '{value.Relationship.State}' but its evidence is contradictory: {string.Join(" ", fidelityErrors)}");
                    }
                }
                ValidateEntryStructure(field, value, "value", value.Relationship.ValueHiddenListEntry, errors);
                ValidateEntryStructure(field, value, "TaxCatchAll", value.Relationship.TaxCatchAllHiddenListEntry, errors);
            }

            return errors;
        }

        public static IReadOnlyList<string> GetFidelityErrors(
            PageFieldValueSnapshot field,
            PageTaxonomyValueSnapshot value)
        {
            var errors = new List<string>();
            var relationship = value?.Relationship;
            Guid termId;
            if (field?.TaxonomyBinding == null || relationship == null || value == null)
            {
                errors.Add("Taxonomy binding, value, or relationship evidence is missing.");
                return errors;
            }
            if (field.TaxonomyBinding.TermStoreId == Guid.Empty
                || field.TaxonomyBinding.BoundTermSetId == Guid.Empty
                || field.TaxonomyBinding.TextFieldId == Guid.Empty)
            {
                errors.Add("The source taxonomy field binding is incomplete.");
            }
            if (!Guid.TryParse(value.TermGuid, out termId) || termId == Guid.Empty)
            {
                errors.Add("The captured taxonomy value has no valid Term GUID.");
            }
            if (value.WssId <= 0)
            {
                errors.Add("The captured taxonomy value has no positive source WssId.");
            }
            if (relationship.State == TaxonomyRelationshipState.Unknown
                || relationship.State == TaxonomyRelationshipState.Conflict)
            {
                errors.Add($"The source taxonomy relationship state is '{relationship.State}'.");
            }

            var valueEntry = relationship.ValueHiddenListEntry;
            if (!EntryIdentifies(valueEntry, field.TaxonomyBinding.TermStoreId, termId)
                || !HasLocalizedIdentity(valueEntry)
                || valueEntry.WssId != value.WssId
                || !string.Equals(valueEntry.PreferredTerm(value.Label), value.Label, StringComparison.Ordinal)
                || string.IsNullOrWhiteSpace(valueEntry.PreferredPath(value.Label)))
            {
                errors.Add("The source value WssId does not resolve to the exact captured hidden-list Term identity, label, and path.");
            }

            switch (relationship.State)
            {
                case TaxonomyRelationshipState.LiveInBoundTermSet:
                    if (valueEntry == null
                        || relationship.LiveTermSetId != field.TaxonomyBinding.BoundTermSetId
                        || !relationship.LiveTermAvailableForTagging.HasValue
                        || valueEntry.TermSetId != field.TaxonomyBinding.BoundTermSetId
                        || !string.Equals(relationship.LiveTermLabel, value.Label, StringComparison.Ordinal)
                        || !string.Equals(relationship.LiveTermPath, valueEntry.PreferredPath(value.Label), StringComparison.Ordinal))
                    {
                        errors.Add("The live Term evidence does not exactly resolve inside the field's bound TermSet.");
                    }
                    break;
                case TaxonomyRelationshipState.LiveOutsideBoundTermSet:
                    if (valueEntry == null
                        || !relationship.LiveTermSetId.HasValue
                        || relationship.LiveTermSetId == Guid.Empty
                        || relationship.LiveTermSetId == field.TaxonomyBinding.BoundTermSetId
                        || !relationship.LiveTermAvailableForTagging.HasValue
                        || !string.Equals(relationship.LiveTermLabel, value.Label, StringComparison.Ordinal)
                        || !string.Equals(relationship.LiveTermPath, valueEntry.PreferredPath(value.Label), StringComparison.Ordinal))
                    {
                        errors.Add("The live outside-bound Term evidence is incomplete or contradictory.");
                    }
                    var catchAll = relationship.TaxCatchAllHiddenListEntry;
                    if (!EntryIdentifies(catchAll, field.TaxonomyBinding.TermStoreId, termId)
                        || !HasLocalizedIdentity(catchAll)
                        || catchAll.TermSetId != field.TaxonomyBinding.BoundTermSetId
                        || !string.Equals(catchAll.CatchAllData, "UNVALIDATED", StringComparison.Ordinal)
                        || !string.Equals(catchAll.PreferredTerm(value.Label), value.Label, StringComparison.Ordinal))
                    {
                        errors.Add("TaxCatchAll does not preserve the exact UNVALIDATED relationship in the bound TermSet.");
                    }
                    if (valueEntry != null
                        && valueEntry.TermSetId != field.TaxonomyBinding.BoundTermSetId
                        && valueEntry.TermSetId != relationship.LiveTermSetId)
                    {
                        errors.Add("The value hidden-list row belongs to neither the bound nor the live TermSet.");
                    }
                    if (valueEntry != null
                        && catchAll != null
                        && valueEntry.TermSetId == field.TaxonomyBinding.BoundTermSetId
                        && !SameEntry(valueEntry, catchAll))
                    {
                        errors.Add("The value and TaxCatchAll identify different source rows for the same bound TermSet relationship.");
                    }
                    break;
                case TaxonomyRelationshipState.DanglingTermAbsent:
                    if (relationship.LiveTermSetId.HasValue)
                    {
                        errors.Add("A dangling Term relationship cannot also identify a live TermSet.");
                    }
                    if (valueEntry != null
                        && valueEntry.TermSetId != field.TaxonomyBinding.BoundTermSetId)
                    {
                        errors.Add("The dangling hidden-list row is not attached to the field's bound TermSet.");
                    }
                    if (valueEntry != null && !PageTaxonomySearchIdentity.IsExact(
                            valueEntry.CatchAllData,
                            field.TaxonomyBinding.TermStoreId,
                            field.TaxonomyBinding.BoundTermSetId,
                            termId))
                    {
                        errors.Add("The dangling hidden-list row has no exact search identity.");
                    }
                    var danglingCatchAll = relationship.TaxCatchAllHiddenListEntry;
                    if (!EntryIdentifies(danglingCatchAll, field.TaxonomyBinding.TermStoreId, termId)
                        || danglingCatchAll.TermSetId != field.TaxonomyBinding.BoundTermSetId
                        || !SameEntry(valueEntry, danglingCatchAll))
                    {
                        errors.Add("TaxCatchAll does not reference the exact dangling value row in the bound TermSet.");
                    }
                    break;
            }

            return errors;
        }

        private static bool EntryIdentifies(TaxonomyHiddenListEntrySnapshot entry, Guid storeId, Guid termId)
        {
            return entry != null
                && entry.WssId > 0
                && entry.TermStoreId == storeId
                && entry.TermId == termId;
        }

        private static bool HasLocalizedIdentity(TaxonomyHiddenListEntrySnapshot entry)
        {
            return entry != null
                && (entry.Terms ?? new List<TaxonomyLocalizedTextSnapshot>())
                    .Any(value => value != null && !string.IsNullOrWhiteSpace(value.Value))
                && (entry.Paths ?? new List<TaxonomyLocalizedTextSnapshot>())
                    .Any(value => value != null && !string.IsNullOrWhiteSpace(value.Value));
        }

        private static bool SameEntry(
            TaxonomyHiddenListEntrySnapshot left,
            TaxonomyHiddenListEntrySnapshot right)
        {
            if (left == null || right == null)
            {
                return false;
            }
            return left.WssId == right.WssId
                && left.TermStoreId == right.TermStoreId
                && left.TermSetId == right.TermSetId
                && left.TermId == right.TermId
                && string.Equals(left.Title, right.Title, StringComparison.Ordinal)
                && string.Equals(left.CatchAllData, right.CatchAllData, StringComparison.Ordinal)
                && string.Equals(left.CatchAllDataLabel, right.CatchAllDataLabel, StringComparison.Ordinal)
                && LocalizedValues(left.Terms).SequenceEqual(LocalizedValues(right.Terms), StringComparer.Ordinal)
                && LocalizedValues(left.Paths).SequenceEqual(LocalizedValues(right.Paths), StringComparer.Ordinal);
        }

        private static IEnumerable<string> LocalizedValues(IEnumerable<TaxonomyLocalizedTextSnapshot> values)
        {
            return (values ?? Array.Empty<TaxonomyLocalizedTextSnapshot>())
                .Where(value => value != null)
                .OrderBy(value => value.FieldInternalName, StringComparer.Ordinal)
                .Select(value => PageTaxonomyRelationshipProof.Encode(value.FieldInternalName) + "|" + PageTaxonomyRelationshipProof.Encode(value.Value));
        }

        private static void ValidateEntryStructure(
            PageFieldValueSnapshot field,
            PageTaxonomyValueSnapshot value,
            string role,
            TaxonomyHiddenListEntrySnapshot entry,
            ICollection<string> errors)
        {
            if (entry == null)
            {
                return;
            }
            if (entry.Terms == null || entry.Paths == null)
            {
                errors.Add($"Field '{field.InternalName}' Term '{value.TermGuid}' has a {role} hidden-list row with a null localized-value collection.");
                return;
            }
            if (entry.Terms.Any(item => item == null || !IsLanguageFieldName(item.FieldInternalName, "Term"))
                || entry.Paths.Any(item => item == null || !IsLanguageFieldName(item.FieldInternalName, "Path"))
                || entry.Terms.GroupBy(item => item.FieldInternalName, StringComparer.OrdinalIgnoreCase).Any(group => group.Count() > 1)
                || entry.Paths.GroupBy(item => item.FieldInternalName, StringComparer.OrdinalIgnoreCase).Any(group => group.Count() > 1))
            {
                errors.Add($"Field '{field.InternalName}' Term '{value.TermGuid}' has malformed or duplicate {role} localized hidden-list evidence.");
            }
        }

        private static bool IsLanguageFieldName(string fieldName, string prefix)
        {
            if (string.IsNullOrWhiteSpace(fieldName)
                || !fieldName.StartsWith(prefix, StringComparison.Ordinal)
                || fieldName.Length == prefix.Length)
            {
                return false;
            }
            int ignored;
            return int.TryParse(fieldName.Substring(prefix.Length), NumberStyles.Integer, CultureInfo.InvariantCulture, out ignored);
        }

    }
}
