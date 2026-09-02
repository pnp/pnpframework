using PnP.Framework.Migration.Packaging;
using PnP.Framework.Migration.Taxonomy;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;

namespace PnP.Framework.Migration.Pages.Fields.Taxonomy
{
    internal static class PageTaxonomyRelationshipProof
    {
        public static void Seal(PageFieldValueSnapshot field)
        {
            if (field == null)
            {
                throw new ArgumentNullException(nameof(field));
            }

            field.TaxonomyValueSetSha256 = ComputeFieldValueSetSha256(field);
            foreach (var value in field.TaxonomyValues ?? new List<PageTaxonomyValueSnapshot>())
            {
                if (value?.Relationship == null)
                {
                    continue;
                }

                value.Relationship.SourceFieldValueSetSha256 = field.TaxonomyValueSetSha256;
                value.Relationship.EvidenceSha256 = ComputeEvidenceSha256(field, value);
            }
        }

        public static string ComputeFieldValueSetSha256(PageFieldValueSnapshot field)
        {
            var lines = new List<string>
            {
                "pnp-page-taxonomy-field-values/v1",
                field.Id.ToString("D"),
                Encode(field.InternalName),
                field.TaxonomyBinding?.TermStoreId.ToString("D") ?? string.Empty,
                field.TaxonomyBinding?.BoundTermSetId.ToString("D") ?? string.Empty,
                field.TaxonomyBinding?.TextFieldId.ToString("D") ?? string.Empty,
                field.TaxonomyBinding?.Open == true ? "true" : "false"
            };
            var values = (field.TaxonomyValues ?? new List<PageTaxonomyValueSnapshot>())
                .Where(value => value != null)
                .OrderBy(value => value.TermGuid, StringComparer.OrdinalIgnoreCase)
                .ThenBy(value => value.WssId)
                .ThenBy(value => value.Label, StringComparer.Ordinal)
                .ToArray();
            lines.Add(values.Length.ToString(CultureInfo.InvariantCulture));
            lines.AddRange(values.Select(value => string.Join("|",
                NormalizeGuid(value.TermGuid),
                value.WssId.ToString(CultureInfo.InvariantCulture),
                Encode(value.Label))));
            return MigrationDigest.ComputeSha256(string.Join("\n", lines));
        }

        public static string ComputeEvidenceSha256(
            PageFieldValueSnapshot field,
            PageTaxonomyValueSnapshot value)
        {
            var relationship = value.Relationship;
            if (relationship == null)
            {
                throw new ArgumentException("A taxonomy relationship is required.", nameof(value));
            }

            var lines = new List<string>
            {
                "pnp-taxonomy-value-relationship-evidence/v1",
                field.Id.ToString("D"),
                Encode(field.InternalName),
                relationship.SourceFieldValueSetSha256 ?? ComputeFieldValueSetSha256(field),
                NormalizeGuid(value.TermGuid),
                value.WssId.ToString(CultureInfo.InvariantCulture),
                Encode(value.Label),
                relationship.CapturedAtUtc.ToUniversalTime().ToString("O", CultureInfo.InvariantCulture),
                relationship.State.ToString(),
                relationship.LiveTermSetId?.ToString("D") ?? string.Empty,
                Encode(relationship.LiveTermSetName),
                Encode(relationship.LiveTermLabel),
                Encode(relationship.LiveTermPath),
                relationship.LiveTermAvailableForTagging.HasValue
                    ? relationship.LiveTermAvailableForTagging.Value ? "true" : "false"
                    : string.Empty
            };
            AddEntry(lines, "value", relationship.ValueHiddenListEntry);
            AddEntry(lines, "tax-catch-all", relationship.TaxCatchAllHiddenListEntry);
            lines.AddRange((relationship.Diagnostics ?? new List<string>())
                .OrderBy(item => item, StringComparer.Ordinal)
                .Select(item => "diagnostic|" + Encode(item)));
            return MigrationDigest.ComputeSha256(string.Join("\n", lines));
        }

        internal static string Encode(string value)
        {
            return Convert.ToBase64String(Encoding.UTF8.GetBytes(value ?? string.Empty));
        }

        private static void AddEntry(
            ICollection<string> lines,
            string role,
            TaxonomyHiddenListEntrySnapshot entry)
        {
            if (entry == null)
            {
                lines.Add(role + "|<null>");
                return;
            }
            lines.Add(string.Join("|",
                role,
                entry.WssId.ToString(CultureInfo.InvariantCulture),
                entry.TermStoreId.ToString("D"),
                entry.TermSetId.ToString("D"),
                entry.TermId.ToString("D"),
                Encode(entry.Title),
                Encode(entry.CatchAllData),
                Encode(entry.CatchAllDataLabel)));
            foreach (var term in (entry.Terms ?? new List<TaxonomyLocalizedTextSnapshot>())
                         .Where(item => item != null)
                         .OrderBy(item => item.FieldInternalName, StringComparer.Ordinal))
            {
                lines.Add(role + "-term|" + Encode(term.FieldInternalName) + "|" + Encode(term.Value));
            }
            foreach (var path in (entry.Paths ?? new List<TaxonomyLocalizedTextSnapshot>())
                         .Where(item => item != null)
                         .OrderBy(item => item.FieldInternalName, StringComparer.Ordinal))
            {
                lines.Add(role + "-path|" + Encode(path.FieldInternalName) + "|" + Encode(path.Value));
            }
        }

        private static string NormalizeGuid(string value)
        {
            Guid parsed;
            return Guid.TryParse(value, out parsed) ? parsed.ToString("D") : value ?? string.Empty;
        }
    }
}
