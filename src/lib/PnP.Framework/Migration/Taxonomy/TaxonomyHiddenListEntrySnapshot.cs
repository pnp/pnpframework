using System;
using System.Collections.Generic;
using System.Linq;

namespace PnP.Framework.Migration.Taxonomy
{
    public sealed class TaxonomyHiddenListEntrySnapshot
    {
        public int WssId { get; set; }

        public Guid TermStoreId { get; set; }

        public Guid TermSetId { get; set; }

        public Guid TermId { get; set; }

        public string Title { get; set; }

        public string CatchAllData { get; set; }

        public string CatchAllDataLabel { get; set; }

        public IList<TaxonomyLocalizedTextSnapshot> Terms { get; set; } = new List<TaxonomyLocalizedTextSnapshot>();

        public IList<TaxonomyLocalizedTextSnapshot> Paths { get; set; } = new List<TaxonomyLocalizedTextSnapshot>();

        public string PreferredTerm(string capturedLabel)
        {
            var terms = Terms ?? new List<TaxonomyLocalizedTextSnapshot>();
            return terms
                .Where(value => value != null && !string.IsNullOrWhiteSpace(value.Value))
                .Select(value => value.Value)
                .FirstOrDefault(value => string.Equals(value, capturedLabel, StringComparison.Ordinal))
                ?? terms.FirstOrDefault(value => value != null && !string.IsNullOrWhiteSpace(value.Value))?.Value
                ?? Title;
        }

        public string PreferredPath(string capturedLabel)
        {
            var terms = Terms ?? new List<TaxonomyLocalizedTextSnapshot>();
            var matchingTermField = terms
                .FirstOrDefault(value => value != null
                    && string.Equals(value.Value, capturedLabel, StringComparison.Ordinal))?
                .FieldInternalName;
            var matchingPathField = string.IsNullOrWhiteSpace(matchingTermField)
                || !matchingTermField.StartsWith("Term", StringComparison.Ordinal)
                || matchingTermField.Length == "Term".Length
                ? null
                : "Path" + matchingTermField.Substring("Term".Length);
            return (Paths ?? new List<TaxonomyLocalizedTextSnapshot>())
                .FirstOrDefault(value => value != null
                    && string.Equals(value.FieldInternalName, matchingPathField, StringComparison.OrdinalIgnoreCase)
                    && !string.IsNullOrWhiteSpace(value.Value))?
                .Value
                ?? (Paths ?? new List<TaxonomyLocalizedTextSnapshot>()).FirstOrDefault(value => value != null && !string.IsNullOrWhiteSpace(value.Value))?.Value
                ?? PreferredTerm(capturedLabel);
        }
    }
}
