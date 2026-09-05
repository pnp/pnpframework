using PnP.Framework.Migration.Lists.Items;
using PnP.Framework.Migration.Lists.Planning;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;

namespace PnP.Framework.Migration.Lists.Execution
{
    internal static class ListItemValueComparer
    {
        public static bool Matches(
            ListItemValueSnapshot source,
            ListItemValueSnapshot actual,
            ListFieldMaterializationPlan fieldPlan,
            IDictionary<Guid, ListMaterializationReceipt> receipts,
            out string mismatch)
        {
            mismatch = null;
            switch (source.Kind)
            {
                case ListItemValueKind.String:
                    return Match(string.Equals(source.ScalarValue ?? string.Empty, actual.ScalarValue ?? actual.RawValue ?? string.Empty, StringComparison.Ordinal), "string", out mismatch);
                case ListItemValueKind.StringCollection:
                    return Match(source.StringValues.SequenceEqual(actual.StringValues, StringComparer.Ordinal), "string collection", out mismatch);
                case ListItemValueKind.Boolean:
                    bool sourceBoolean;
                    bool actualBoolean;
                    return Match(bool.TryParse(source.ScalarValue, out sourceBoolean)
                        && bool.TryParse(actual.ScalarValue, out actualBoolean)
                        && sourceBoolean == actualBoolean, "Boolean", out mismatch);
                case ListItemValueKind.Number:
                    decimal sourceNumber;
                    decimal actualNumber;
                    return Match(decimal.TryParse(source.ScalarValue, NumberStyles.Any, CultureInfo.InvariantCulture, out sourceNumber)
                        && decimal.TryParse(actual.ScalarValue, NumberStyles.Any, CultureInfo.InvariantCulture, out actualNumber)
                        && sourceNumber == actualNumber, "number", out mismatch);
                case ListItemValueKind.DateTime:
                    DateTimeOffset sourceDate;
                    DateTimeOffset actualDate;
                    return Match(DateTimeOffset.TryParse(source.ScalarValue, CultureInfo.InvariantCulture, DateTimeStyles.RoundtripKind, out sourceDate)
                        && DateTimeOffset.TryParse(actual.ScalarValue, CultureInfo.InvariantCulture, DateTimeStyles.RoundtripKind, out actualDate)
                        && sourceDate.ToUniversalTime() == actualDate.ToUniversalTime(), "date/time", out mismatch);
                case ListItemValueKind.Guid:
                    Guid sourceGuid;
                    Guid actualGuid;
                    return Match(Guid.TryParse(source.ScalarValue, out sourceGuid)
                        && Guid.TryParse(actual.ScalarValue, out actualGuid)
                        && sourceGuid == actualGuid, "GUID", out mismatch);
                case ListItemValueKind.Url:
                    return Match(source.UrlValue != null && actual.UrlValue != null
                        && string.Equals(source.UrlValue.Url, actual.UrlValue.Url, StringComparison.Ordinal)
                        && string.Equals(source.UrlValue.Description ?? string.Empty, actual.UrlValue.Description ?? string.Empty, StringComparison.Ordinal), "URL", out mismatch);
                case ListItemValueKind.Lookup:
                case ListItemValueKind.LookupCollection:
                    ListMaterializationReceipt lookupReceipt;
                    if (!fieldPlan.SourceLookupListId.HasValue || !receipts.TryGetValue(fieldPlan.SourceLookupListId.Value, out lookupReceipt))
                    {
                        mismatch = "lookup mapping receipt is unavailable";
                        return false;
                    }
                    var expectedLookupIds = source.LookupValues.Select(value =>
                    {
                        int mapped;
                        return lookupReceipt.TargetItemIds.TryGetValue(value.LookupId, out mapped) ? mapped : -1;
                    }).ToArray();
                    return Match(expectedLookupIds.All(value => value > 0)
                        && expectedLookupIds.SequenceEqual(actual.LookupValues.Select(value => value.LookupId)), "lookup target IDs", out mismatch);
                case ListItemValueKind.Taxonomy:
                case ListItemValueKind.TaxonomyCollection:
                    return Match(source.TaxonomyValues.Select(value => (value.TermGuid ?? string.Empty).ToUpperInvariant())
                        .SequenceEqual(actual.TaxonomyValues.Select(value => (value.TermGuid ?? string.Empty).ToUpperInvariant()), StringComparer.Ordinal)
                        && source.TaxonomyValues.Select(value => value.Label ?? string.Empty)
                            .SequenceEqual(actual.TaxonomyValues.Select(value => value.Label ?? string.Empty), StringComparer.Ordinal), "taxonomy Term GUIDs/labels", out mismatch);
                default:
                    mismatch = "value kind " + source.Kind + " has no approved semantic verifier";
                    return false;
            }
        }

        private static bool Match(bool value, string subject, out string mismatch)
        {
            mismatch = value ? null : subject + " does not match";
            return value;
        }
    }
}
