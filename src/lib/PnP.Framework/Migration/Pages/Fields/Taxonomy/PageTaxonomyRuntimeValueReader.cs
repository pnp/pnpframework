using System;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text.Json;

namespace PnP.Framework.Migration.Pages.Fields.Taxonomy
{
    internal sealed class PageTaxonomyRuntimeValue
    {
        public bool IsCollection { get; set; }

        public IList<PageTaxonomyValueSnapshot> Values { get; set; } = new List<PageTaxonomyValueSnapshot>();
    }

    internal static class PageTaxonomyRuntimeValueReader
    {
        private const string TaxonomyValueType = "SP.Taxonomy.TaxonomyFieldValue";
        private const string TaxonomyCollectionType = "SP.Taxonomy.TaxonomyFieldValueCollection";

        public static bool TryRead(object value, out PageTaxonomyRuntimeValue result)
        {
            result = null;
            if (!TryReadDictionary(value, out var dictionary))
            {
                return false;
            }

            var objectType = ReadString(dictionary, "_ObjectType_");
            if (string.Equals(objectType, TaxonomyCollectionType, StringComparison.OrdinalIgnoreCase))
            {
                result = new PageTaxonomyRuntimeValue { IsCollection = true };
                if (TryGetValue(dictionary, "_Child_Items_", out var children))
                {
                    result.Values = Enumerate(children)
                        .Select(ReadValue)
                        .Where(item => item != null)
                        .ToList();
                }
                return true;
            }

            if (!string.Equals(objectType, TaxonomyValueType, StringComparison.OrdinalIgnoreCase))
            {
                return false;
            }

            result = new PageTaxonomyRuntimeValue
            {
                IsCollection = false,
                Values = new List<PageTaxonomyValueSnapshot> { ReadValue(dictionary) }
            };
            return true;
        }

        private static PageTaxonomyValueSnapshot ReadValue(object value)
        {
            return TryReadDictionary(value, out var dictionary)
                ? ReadValue(dictionary)
                : null;
        }

        private static PageTaxonomyValueSnapshot ReadValue(IDictionary<string, object> dictionary)
        {
            return new PageTaxonomyValueSnapshot
            {
                Label = ReadString(dictionary, "Label"),
                TermGuid = ReadString(dictionary, "TermGuid"),
                WssId = ReadInt32(dictionary, "WssId")
            };
        }

        private static IEnumerable<object> Enumerate(object value)
        {
            if (value is JsonElement element && element.ValueKind == JsonValueKind.Array)
            {
                return element.EnumerateArray().Select(item => (object)item).ToArray();
            }
            if (value is IEnumerable enumerable && !(value is string))
            {
                return enumerable.Cast<object>();
            }
            return Array.Empty<object>();
        }

        private static bool TryReadDictionary(object value, out IDictionary<string, object> result)
        {
            result = null;
            if (value is IDictionary<string, object> dictionary)
            {
                result = new Dictionary<string, object>(dictionary, StringComparer.OrdinalIgnoreCase);
                return true;
            }
            if (value is IReadOnlyDictionary<string, object> readOnlyDictionary)
            {
                result = readOnlyDictionary.ToDictionary(
                    item => item.Key,
                    item => item.Value,
                    StringComparer.OrdinalIgnoreCase);
                return true;
            }
            if (value is IDictionary untypedDictionary)
            {
                var normalized = new Dictionary<string, object>(StringComparer.OrdinalIgnoreCase);
                foreach (DictionaryEntry entry in untypedDictionary)
                {
                    if (entry.Key != null)
                    {
                        normalized[Convert.ToString(entry.Key, CultureInfo.InvariantCulture)] = entry.Value;
                    }
                }
                result = normalized;
                return true;
            }
            if (value is JsonElement element && element.ValueKind == JsonValueKind.Object)
            {
                result = element.EnumerateObject().ToDictionary(
                    property => property.Name,
                    property => (object)property.Value,
                    StringComparer.OrdinalIgnoreCase);
                return true;
            }
            return false;
        }

        private static bool TryGetValue(
            IDictionary<string, object> dictionary,
            string key,
            out object value)
        {
            return dictionary.TryGetValue(key, out value);
        }

        private static string ReadString(IDictionary<string, object> dictionary, string key)
        {
            if (!TryGetValue(dictionary, key, out var value) || value == null)
            {
                return null;
            }
            if (value is JsonElement element)
            {
                if (element.ValueKind == JsonValueKind.String)
                {
                    return element.GetString();
                }
                return element.ValueKind == JsonValueKind.Null
                    ? null
                    : element.GetRawText();
            }
            return Convert.ToString(value, CultureInfo.InvariantCulture);
        }

        private static int ReadInt32(IDictionary<string, object> dictionary, string key)
        {
            if (!TryGetValue(dictionary, key, out var value) || value == null)
            {
                return 0;
            }
            if (value is JsonElement element)
            {
                return element.ValueKind == JsonValueKind.Number && element.TryGetInt32(out var number)
                    ? number
                    : int.TryParse(element.ToString(), NumberStyles.Integer, CultureInfo.InvariantCulture, out number)
                        ? number
                        : 0;
            }
            return value is int integer
                ? integer
                : int.TryParse(Convert.ToString(value, CultureInfo.InvariantCulture), NumberStyles.Integer, CultureInfo.InvariantCulture, out var parsed)
                    ? parsed
                    : 0;
        }
    }
}
