using PnP.Framework.Migration.Diagnostics;
using PnP.Framework.Migration.Evidence;
using PnP.Framework.Migration.Lists.Items;
using PnP.Framework.Migration.Schema.Fields;
using System;
using System.Collections.Generic;
using System.Linq;

namespace PnP.Framework.Migration.Pages.Publishing.Reporting.Sections
{
    internal static class MigrationReportSectionFormatter
    {
        public static string[] Row(params object[] values) => PublishingPageReportValueFormatter.Row(values);

        public static string Format(object value) => PublishingPageReportValueFormatter.Format(value);

        public static string Join(IEnumerable<string> values) => PublishingPageReportValueFormatter.Join(values);

        public static string Summarize(string value) => PublishingPageReportValueFormatter.SummarizePayload(value);

        public static void AppendIssues(
            MarkdownReportWriter writer,
            string heading,
            IEnumerable<MigrationIssue> issues,
            int headingLevel = 2)
        {
            if (!string.IsNullOrWhiteSpace(heading))
            {
                writer.Heading(headingLevel, heading);
            }
            writer.Table(null,
                new[] { "Code", "Severity", "Ingredient", "Subject", "Source", "Target", "Message" },
                (issues ?? Enumerable.Empty<MigrationIssue>()).Select(value => Row(
                    value.Code,
                    value.Severity,
                    value.Ingredient,
                    value.Subject,
                    value.SourceIdentity,
                    value.TargetIdentity,
                    value.Message)));
        }

        public static string IssueSummary(IEnumerable<MigrationIssue> issues)
        {
            return Join((issues ?? Enumerable.Empty<MigrationIssue>()).Select(value =>
                $"{value.Code} ({value.Severity}): {value.Message}"));
        }

        public static string FormatTaxonomy(TaxonomyFieldBindingSnapshot value)
        {
            return value == null
                ? null
                : $"store={value.SourceTermStoreId:D}; set={value.SourceTermSetId:D}; hiddenText={value.HiddenTextFieldId:D}; open={value.Open}";
        }

        public static string FormatEvidenceSource(EvidenceSource value)
        {
            return value == null
                ? null
                : $"exchange={Format(value.ExchangeId)}; payloadSha256={Format(value.PayloadSha256)}; selector={Format(value.Selector)}";
        }

        public static string FormatArtifact(ListBinaryArtifactSnapshot value)
        {
            if (value == null)
            {
                return null;
            }
            var artifact = value.Artifact;
            return $"availability={value.Availability}; sha256={Format(artifact?.Sha256)}; bytes={Format(artifact?.Length)}; mediaType={Format(artifact?.MediaType)}; originalName={Format(artifact?.OriginalName)}; inlineBase64={Summarize(value.ContentBase64)}; diagnostics={Join(value.Diagnostics)}";
        }

        public static string SummarizeListItemValue(ListItemValueSnapshot value)
        {
            switch (value.Kind)
            {
                case ListItemValueKind.Url:
                    return value.UrlValue == null
                        ? null
                        : $"url={Format(value.UrlValue.Url)}; description={Format(value.UrlValue.Description)}";
                case ListItemValueKind.StringCollection:
                    return Join(value.StringValues.Select(Summarize));
                case ListItemValueKind.User:
                case ListItemValueKind.UserCollection:
                case ListItemValueKind.Lookup:
                case ListItemValueKind.LookupCollection:
                    return Join(value.LookupValues.Select(item => $"id={item.LookupId}; value={Format(item.LookupValue)}"));
                case ListItemValueKind.Taxonomy:
                case ListItemValueKind.TaxonomyCollection:
                    return Join(value.TaxonomyValues.Select(item => $"label={Format(item.Label)}; termGuid={Format(item.TermGuid)}; sourceWssId={item.WssId}"));
                case ListItemValueKind.ByteArray:
                case ListItemValueKind.Unsupported:
                    return Summarize(value.ScalarValue ?? value.RawValueJson ?? value.RawValue);
                default:
                    return Summarize(value.ScalarValue);
            }
        }
    }
}
