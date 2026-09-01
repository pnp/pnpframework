using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.Taxonomy;
using PnP.Framework.Migration.Evidence;
using System;
using System.Globalization;
using System.Linq;
using System.Text.Json;

namespace PnP.Framework.Migration.Lists.Items
{
    internal static class ListItemValueSerializer
    {
        public static ListItemValueSnapshot Serialize(string internalName, object value)
        {
            var result = new ListItemValueSnapshot
            {
                InternalName = internalName,
                RawType = value == null ? null : value.GetType().FullName,
                RawValue = SafeConvertToString(value)
            };
            if (value == null)
            {
                result.Kind = ListItemValueKind.Null;
                return result;
            }

            if (value is FieldUrlValue)
            {
                var url = (FieldUrlValue)value;
                result.Kind = ListItemValueKind.Url;
                result.UrlValue = new ListItemUrlValueSnapshot { Url = url.Url, Description = url.Description };
            }
            else if (value is TaxonomyFieldValueCollection)
            {
                result.Kind = ListItemValueKind.TaxonomyCollection;
                result.TaxonomyValues = ((TaxonomyFieldValueCollection)value).Select(ToTaxonomy).ToList();
            }
            else if (value is TaxonomyFieldValue)
            {
                result.Kind = ListItemValueKind.Taxonomy;
                result.TaxonomyValues.Add(ToTaxonomy((TaxonomyFieldValue)value));
            }
            else if (value is FieldUserValue[])
            {
                result.Kind = ListItemValueKind.UserCollection;
                result.LookupValues = ((FieldUserValue[])value).Select(ToLookup).ToList();
            }
            else if (value is FieldUserValue)
            {
                result.Kind = ListItemValueKind.User;
                result.LookupValues.Add(ToLookup((FieldUserValue)value));
            }
            else if (value is FieldLookupValue[])
            {
                result.Kind = ListItemValueKind.LookupCollection;
                result.LookupValues = ((FieldLookupValue[])value).Select(ToLookup).ToList();
            }
            else if (value is FieldLookupValue)
            {
                result.Kind = ListItemValueKind.Lookup;
                result.LookupValues.Add(ToLookup((FieldLookupValue)value));
            }
            else if (value is DateTime)
            {
                result.Kind = ListItemValueKind.DateTime;
                result.ScalarValue = ((DateTime)value).ToUniversalTime().ToString("O", CultureInfo.InvariantCulture);
            }
            else if (value is bool)
            {
                result.Kind = ListItemValueKind.Boolean;
                result.ScalarValue = (bool)value ? "true" : "false";
            }
            else if (IsNumber(value))
            {
                result.Kind = ListItemValueKind.Number;
                result.ScalarValue = Convert.ToString(value, CultureInfo.InvariantCulture);
            }
            else if (value is Guid)
            {
                result.Kind = ListItemValueKind.Guid;
                result.ScalarValue = ((Guid)value).ToString("D");
            }
            else if (value is byte[])
            {
                result.Kind = ListItemValueKind.ByteArray;
                result.ScalarValue = Convert.ToBase64String((byte[])value);
            }
            else if (value is string[])
            {
                result.Kind = ListItemValueKind.StringCollection;
                result.StringValues = ((string[])value).ToList();
            }
            else if (value is string)
            {
                result.Kind = ListItemValueKind.String;
                result.ScalarValue = (string)value;
            }
            else
            {
                result.Kind = ListItemValueKind.Unsupported;
                result.Availability = EvidenceAvailability.Partial;
                result.Diagnostics.Add("No typed list-item serializer is registered for this runtime value. Raw evidence is retained.");
            }

            try
            {
                result.RawValueJson = JsonSerializer.Serialize(value, value.GetType());
            }
            catch (Exception exception) when (exception is JsonException || exception is NotSupportedException)
            {
                result.Diagnostics.Add("Best-effort raw JSON serialization was unavailable: " + exception.Message);
            }
            return result;
        }

        private static bool IsNumber(object value)
        {
            return value is byte || value is short || value is int || value is long || value is float || value is double || value is decimal;
        }

        private static ListItemLookupValueSnapshot ToLookup(FieldLookupValue value)
        {
            return new ListItemLookupValueSnapshot { LookupId = value.LookupId, LookupValue = value.LookupValue };
        }

        private static ListItemTaxonomyValueSnapshot ToTaxonomy(TaxonomyFieldValue value)
        {
            return new ListItemTaxonomyValueSnapshot { Label = value.Label, TermGuid = value.TermGuid, WssId = value.WssId };
        }

        private static string SafeConvertToString(object value)
        {
            try
            {
                return value == null ? null : Convert.ToString(value, CultureInfo.InvariantCulture);
            }
            catch (Exception)
            {
                return null;
            }
        }
    }
}
