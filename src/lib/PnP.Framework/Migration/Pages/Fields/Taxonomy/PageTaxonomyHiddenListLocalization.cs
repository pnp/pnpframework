using PnP.Framework.Migration.Taxonomy;
using System;
using System.Collections.Generic;
using System.Linq;

namespace PnP.Framework.Migration.Pages.Fields.Taxonomy
{
    internal static class PageTaxonomyHiddenListLocalization
    {
        public static IReadOnlyList<string> GetTargetCoverageErrors(
            TaxonomyHiddenListEntrySnapshot source,
            IEnumerable<string> targetTermFields,
            IEnumerable<string> targetPathFields)
        {
            if (source == null)
            {
                return new[] { "The sealed source TaxonomyHiddenList row is unavailable." };
            }

            var targetTerms = new HashSet<string>(targetTermFields ?? Array.Empty<string>(), StringComparer.OrdinalIgnoreCase);
            var targetPaths = new HashSet<string>(targetPathFields ?? Array.Empty<string>(), StringComparer.OrdinalIgnoreCase);
            var errors = new List<string>();
            foreach (var fieldName in (source.Terms ?? new List<TaxonomyLocalizedTextSnapshot>())
                         .Where(value => value != null && !string.IsNullOrWhiteSpace(value.FieldInternalName))
                         .Select(value => value.FieldInternalName)
                         .Distinct(StringComparer.OrdinalIgnoreCase)
                         .Where(value => !targetTerms.Contains(value)))
            {
                errors.Add($"The target TaxonomyHiddenList is missing captured localized field '{fieldName}'.");
            }
            foreach (var fieldName in (source.Paths ?? new List<TaxonomyLocalizedTextSnapshot>())
                         .Where(value => value != null && !string.IsNullOrWhiteSpace(value.FieldInternalName))
                         .Select(value => value.FieldInternalName)
                         .Distinct(StringComparer.OrdinalIgnoreCase)
                         .Where(value => !targetPaths.Contains(value)))
            {
                errors.Add($"The target TaxonomyHiddenList is missing captured localized field '{fieldName}'.");
            }
            return errors;
        }

        public static bool MatchesCapturedValues(
            TaxonomyHiddenListEntrySnapshot source,
            IEnumerable<string> targetTermFields,
            IEnumerable<string> targetPathFields,
            Func<string, string> readTargetValue)
        {
            if (source == null || readTargetValue == null)
            {
                return false;
            }

            var targetTerms = new HashSet<string>(targetTermFields ?? Array.Empty<string>(), StringComparer.OrdinalIgnoreCase);
            var targetPaths = new HashSet<string>(targetPathFields ?? Array.Empty<string>(), StringComparer.OrdinalIgnoreCase);
            return (source.Terms ?? new List<TaxonomyLocalizedTextSnapshot>()).All(value => value != null
                    && targetTerms.Contains(value.FieldInternalName)
                    && string.Equals(readTargetValue(value.FieldInternalName), value.Value ?? string.Empty, StringComparison.Ordinal))
                && (source.Paths ?? new List<TaxonomyLocalizedTextSnapshot>()).All(value => value != null
                    && targetPaths.Contains(value.FieldInternalName)
                    && string.Equals(readTargetValue(value.FieldInternalName), value.Value ?? string.Empty, StringComparison.Ordinal));
        }
    }
}
