using PnP.Framework.Migration.Pages.Publishing.Layouts;
using System.Collections.Generic;
using System.Linq;

namespace PnP.Framework.Migration.Pages.Publishing.Reporting
{
    internal static class PublishingPageLayoutSnapshotReportSection
    {
        public static void Append(MarkdownReportWriter writer, PublishingPageLayoutSnapshot layout)
        {
            writer.Table("Source Page Layout evidence", new[] { "JSON path", "Value", "How to read it" }, new[]
            {
                Row("snapshot.layout.schemaVersion", layout.SchemaVersion, "Version of the nested Page Layout evidence contract."),
                Row("snapshot.layout.evidenceState", layout.EvidenceState, "Readable means exact layout bytes and parsed ingredients were captured; every other value is explicit incomplete evidence."),
                Row("snapshot.layout.availability", layout.Availability, "Captured is required for custom-layout materialization."),
                Row("snapshot.layout.url", layout.Url, "Absolute source layout URL."),
                Row("snapshot.layout.serverRelativeUrl", layout.ServerRelativeUrl, "Exact source master-page-gallery path."),
                Row("snapshot.layout.ownerSiteCollectionUrl", layout.OwnerSiteCollectionUrl, "Site Collection that owns the captured master-page-gallery item; it can differ from the page Site Collection."),
                Row("snapshot.layout.externalToPageSiteCollection", layout.ExternalToPageSiteCollection, "True means Page Layout bytes and schema were captured from an external owner while page-runtime tokens retain the page Web/Site Collection context."),
                Row("snapshot.layout.authorizationEvidence", Authorization(layout), "Present only when the Page Layout owner request returned literal wire HTTP 401/403."),
                Row("snapshot.layout.description", layout.Description, "Source gallery description."),
                Row("snapshot.layout.fileUniqueId", layout.FileUniqueId, "Source layout file identity evidence; it is not reused as the target file ID."),
                Row("snapshot.layout.customizedPageStatus", layout.CustomizedPageStatus, "1 identifies an uncustomized ghosted layout. Only uncustomized reviewed EnterpriseWiki.aspx is reused as target stock."),
                Row("snapshot.layout.setupPath", layout.SetupPath, "SharePoint setup-file provenance when available."),
                Row("snapshot.layout.fileName", layout.FileName, "Source layout file name."),
                Row("snapshot.layout.itemContentTypeId", layout.ItemContentTypeId, "Content type of the layout gallery item, normally System Page Layout."),
                Row("snapshot.layout.associatedContentTypeName", layout.AssociatedContentTypeName, "Publishing page content type selected by the layout."),
                Row("snapshot.layout.associatedContentTypeId", layout.AssociatedContentTypeId, "Exact source associated content type ID."),
                Row("snapshot.layout.title", layout.Title, "Source gallery item title."),
                Row("snapshot.layout.level", layout.Level, "Source layout file publication level evidence."),
                Row("snapshot.layout.checkOutType", layout.CheckOutType, "Source layout checkout evidence."),
                Row("snapshot.layout.versionLabel", layout.VersionLabel, "Captured source layout version label."),
                Row("snapshot.layout.bytes", PublishingPageArtifactReportFormatter.Artifact(layout.Bytes), "Content-addressed exact ASPX bytes. Inline Base64 may be omitted when an artifact store is used."),
                Row("snapshot.layout.contentBase64", Summarize(layout.ContentBase64), "Inline exact ASPX bytes, when the package is self-contained."),
                Row("snapshot.layout.diagnostics", Join(layout.Diagnostics), "Capture diagnostics; absence means no diagnostic was recorded.")
            });

            writer.Table("Page Layout server-control registrations",
                new[] { "Tag prefix", "Namespace", "Assembly", "Interpretation" },
                layout.Registrations.Select(item => Row(
                    item.TagPrefix,
                    item.Namespace,
                    item.Assembly,
                    "Platform registrations are admitted; non-platform assemblies block until a reviewed target capability exists.")));

            writer.Table("Page Layout controls",
                new[] { "Tag prefix", "Control", "ID", "Field name", "Interpretation" },
                layout.Controls.Select(item => Row(
                    item.TagPrefix,
                    item.ControlName,
                    item.Id,
                    item.FieldName,
                    string.IsNullOrWhiteSpace(item.FieldName)
                        ? "A parsed server control with no page-field binding."
                        : "This field belongs to the minimal schema closure required by the layout.")));

            writer.Table("Page Layout Web Part zones",
                new[] { "ID", "Title", "Interpretation" },
                layout.Zones.Select(item => Row(
                    item.Id,
                    item.Title,
                    "A named placement surface defined by the ASPX layout.")));

            writer.Table("Page Layout parsed resource references",
                new[] { "Attribute", "Authored value", "Interpretation" },
                layout.ResourceReferences.Select(item => Row(
                    item.Attribute,
                    item.Value,
                    "Each parsed reference must have exactly one evidence record and one custom-layout plan action.")));

            writer.Table("Page Layout resource evidence",
                new[] { "Attribute", "Authored value", "Evidence", "Resolved source URL", "Artifact", "Inline bytes", "Sources", "Diagnostics" },
                layout.ResourceArtifacts.Select(item => Row(
                    item.Reference.Attribute,
                    item.Reference.Value,
                    item.EvidenceState,
                    item.ResolvedSourceUrl,
                    PublishingPageArtifactReportFormatter.Artifact(item.Artifact),
                    Summarize(item.ContentBase64),
                    PublishingPageArtifactReportFormatter.Sources(item.Sources),
                    Join(item.Diagnostics))));

            PublishingPageSchemaReportSection.AppendSnapshot(writer, layout.AssociatedContentTypeSchema);
        }

        private static string[] Row(params object[] values) => PublishingPageReportValueFormatter.Row(values);

        private static string Join(IEnumerable<string> values) => PublishingPageReportValueFormatter.Join(values);

        private static string Summarize(string value) => PublishingPageReportValueFormatter.SummarizePayload(value);

        private static string Authorization(PublishingPageLayoutSnapshot layout)
        {
            var evidence = layout.AuthorizationEvidence;
            return evidence == null
                ? null
                : $"{evidence.Operation}; HTTP {evidence.HttpStatusCode}; {evidence.RequestUri}; {evidence.ObservedAtUtc:O}; sha256={evidence.EvidenceSha256}";
        }
    }
}
