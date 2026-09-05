using PnP.Framework.Migration.Pages.Fields;
using PnP.Framework.Migration.Pages.Publishing.Packaging;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;

namespace PnP.Framework.Migration.Pages.Publishing.Reporting
{
    internal static class PublishingPageReportValueFormatter
    {
        public static string Format(object value)
        {
            if (value == null)
            {
                return "(null)";
            }

            if (value is DateTime dateTime)
            {
                return dateTime.ToUniversalTime().ToString("O", CultureInfo.InvariantCulture);
            }

            if (value is DateTimeOffset dateTimeOffset)
            {
                return dateTimeOffset.ToUniversalTime().ToString("O", CultureInfo.InvariantCulture);
            }

            var result = Convert.ToString(value, CultureInfo.InvariantCulture);
            return string.IsNullOrEmpty(result) ? "(empty)" : result;
        }

        public static string[] Row(params object[] values)
        {
            return values.Select(Format).ToArray();
        }

        public static string Join(IEnumerable<string> values)
        {
            var items = (values ?? Array.Empty<string>()).Where(value => value != null).ToArray();
            return items.Length == 0 ? null : string.Join("; ", items);
        }

        public static string SummarizePayload(string value)
        {
            if (value == null)
            {
                return null;
            }

            var preview = value.Replace("\r", " ").Replace("\n", " ");
            if (preview.Length > 160)
            {
                preview = preview.Substring(0, 160) + "…";
            }

            return $"length={value.Length}; sha256={PublishingPageDigest.ComputeSha256(value)}; preview={preview}";
        }

        public static string SummarizeFieldValue(PageFieldValueSnapshot field)
        {
            switch (field.Kind)
            {
                case PageFieldValueKind.Url:
                    return field.UrlValue == null
                        ? null
                        : $"url={Format(field.UrlValue.Url)}; description={Format(field.UrlValue.Description)}";
                case PageFieldValueKind.StringCollection:
                    return Join(field.StringValues);
                case PageFieldValueKind.User:
                case PageFieldValueKind.UserCollection:
                case PageFieldValueKind.Lookup:
                case PageFieldValueKind.LookupCollection:
                    return Join(field.LookupValues.Select(value => $"{value.LookupId}:{value.LookupValue}"));
                case PageFieldValueKind.Taxonomy:
                case PageFieldValueKind.TaxonomyCollection:
                    return Join(field.TaxonomyValues.Select(value => $"{value.Label}|{value.TermGuid}|{value.WssId}"));
                case PageFieldValueKind.ByteArray:
                    return SummarizePayload(field.BinaryBase64);
                case PageFieldValueKind.Unsupported:
                    return SummarizePayload(field.RawValueJson ?? field.RawValue);
                default:
                    return SummarizePayload(field.Value);
            }
        }

        public static string EscapeHeading(string value)
        {
            return (value ?? "(unnamed)").Replace("#", "\\#").Replace("`", "\\`");
        }
    }
}
