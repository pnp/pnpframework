using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.Taxonomy;
using PnP.Framework.Migration.Pages.Capture;
using PnP.Framework.Migration.Pages.Fields.Taxonomy;
using System;
using System.Globalization;
using System.Linq;
using System.Text.Json;

namespace PnP.Framework.Migration.Pages.Fields
{
    internal static class PageFieldValueSerializer
    {
        public static PageFieldValueSnapshot Serialize(Field field, object value)
        {
            var snapshot = Create(field, PageCaptureStatus.Captured, PageFieldValueKind.Null);
            snapshot.HasValue = value != null;
            snapshot.RawType = value?.GetType().FullName;
            snapshot.RawValue = SafeConvertToString(value);
            snapshot.RawValueJson = TrySerializeRawValue(value, snapshot.Diagnostics);
            if (value == null)
            {
                return snapshot;
            }

            if (value is FieldUrlValue url)
            {
                snapshot.Kind = PageFieldValueKind.Url;
                snapshot.UrlValue = new PageUrlValueSnapshot
                {
                    Url = url.Url,
                    Description = url.Description
                };
            }
            else if (value is TaxonomyFieldValueCollection taxonomyCollection)
            {
                snapshot.Kind = PageFieldValueKind.TaxonomyCollection;
                snapshot.TaxonomyValues = taxonomyCollection.Select(ToTaxonomyValue).ToList();
            }
            else if (value is TaxonomyFieldValue taxonomy)
            {
                snapshot.Kind = PageFieldValueKind.Taxonomy;
                snapshot.TaxonomyValues.Add(ToTaxonomyValue(taxonomy));
            }
            else if (PageTaxonomyRuntimeValueReader.TryRead(value, out var runtimeTaxonomy))
            {
                snapshot.Kind = runtimeTaxonomy.IsCollection
                    ? PageFieldValueKind.TaxonomyCollection
                    : PageFieldValueKind.Taxonomy;
                snapshot.TaxonomyValues = runtimeTaxonomy.Values.ToList();
            }
            else if (value is FieldUserValue[] users)
            {
                snapshot.Kind = PageFieldValueKind.UserCollection;
                snapshot.LookupValues = users.Select(ToLookupValue).ToList();
            }
            else if (value is FieldUserValue user)
            {
                snapshot.Kind = PageFieldValueKind.User;
                snapshot.LookupValues.Add(ToLookupValue(user));
            }
            else if (value is FieldLookupValue[] lookups)
            {
                snapshot.Kind = PageFieldValueKind.LookupCollection;
                snapshot.LookupValues = lookups.Select(ToLookupValue).ToList();
            }
            else if (value is FieldLookupValue lookup)
            {
                snapshot.Kind = PageFieldValueKind.Lookup;
                snapshot.LookupValues.Add(ToLookupValue(lookup));
            }
            else if (value is DateTime dateTime)
            {
                snapshot.Kind = PageFieldValueKind.DateTime;
                snapshot.Value = dateTime.ToUniversalTime().ToString("O", CultureInfo.InvariantCulture);
            }
            else if (value is bool boolean)
            {
                snapshot.Kind = PageFieldValueKind.Boolean;
                snapshot.Value = boolean ? "true" : "false";
            }
            else if (value is byte || value is short || value is int || value is long || value is float || value is double || value is decimal)
            {
                snapshot.Kind = PageFieldValueKind.Number;
                snapshot.Value = Convert.ToString(value, CultureInfo.InvariantCulture);
            }
            else if (value is Guid guid)
            {
                snapshot.Kind = PageFieldValueKind.Guid;
                snapshot.Value = guid.ToString("D");
            }
            else if (value is byte[] bytes)
            {
                snapshot.Kind = PageFieldValueKind.ByteArray;
                snapshot.BinaryBase64 = Convert.ToBase64String(bytes);
            }
            else if (value is string[] strings)
            {
                snapshot.Kind = PageFieldValueKind.StringCollection;
                snapshot.StringValues = strings.ToList();
            }
            else if (value is string text)
            {
                snapshot.Kind = PageFieldValueKind.String;
                snapshot.Value = text;
            }
            else
            {
                snapshot.Kind = PageFieldValueKind.Unsupported;
                snapshot.CaptureStatus = PageCaptureStatus.CapturedWithLimitations;
                snapshot.Diagnostics.Add("No typed importer is registered for this runtime value. Raw type, text, and best-effort JSON are retained for future recovery.");
            }

            return snapshot;
        }

        public static PageFieldValueSnapshot Create(Field field, PageCaptureStatus captureStatus, PageFieldValueKind kind)
        {
            return new PageFieldValueSnapshot
            {
                Id = field.Id,
                InternalName = field.InternalName,
                Title = field.Title,
                TypeAsString = field.TypeAsString,
                SchemaXml = field.SchemaXml,
                ReadOnly = field.ReadOnlyField,
                Hidden = field.Hidden,
                Required = field.Required,
                Kind = kind,
                CaptureStatus = captureStatus
            };
        }

        public static string SafeConvertToString(object value)
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

        private static PageLookupValueSnapshot ToLookupValue(FieldLookupValue value)
        {
            return new PageLookupValueSnapshot
            {
                LookupId = value.LookupId,
                LookupValue = value.LookupValue
            };
        }

        private static PageTaxonomyValueSnapshot ToTaxonomyValue(TaxonomyFieldValue value)
        {
            return new PageTaxonomyValueSnapshot
            {
                Label = value.Label,
                TermGuid = value.TermGuid,
                WssId = value.WssId
            };
        }

        private static string TrySerializeRawValue(object value, System.Collections.Generic.ICollection<string> diagnostics)
        {
            if (value == null)
            {
                return null;
            }

            try
            {
                return JsonSerializer.Serialize(value, value.GetType());
            }
            catch (Exception exception) when (exception is NotSupportedException || exception is JsonException)
            {
                diagnostics.Add($"Best-effort raw JSON serialization was unavailable: {exception.Message}");
                return null;
            }
        }
    }
}
