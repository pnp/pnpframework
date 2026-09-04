using PnP.Framework.Migration.Lists.Capture;
using PnP.Framework.Migration.Lists.Items;
using System;
using System.Linq;
using static PnP.Framework.Migration.Pages.Publishing.Reporting.Sections.MigrationReportSectionFormatter;

namespace PnP.Framework.Migration.Pages.Publishing.Reporting.Sections
{
    internal static class ListItemMigrationReportSection
    {
        public static void Append(MarkdownReportWriter writer, ListDependencySnapshot source)
        {
            writer.Heading(4, $"Current items, folders, files, and attachments ({source.Items.Count})");
            writer.Paragraph("This is current-state evidence only. Source item IDs are mapped to target-generated IDs; version history and Created/Modified/Author/Editor replay are outside the current contract.");
            writer.Table(null, new[] { "Source item ID", "Source unique ID", "Availability", "Document object", "Values", "Attachments", "Diagnostics" },
                source.Items.OrderBy(value => value.SourceItemId).Select(value => Row(
                    value.SourceItemId,
                    value.SourceUniqueId,
                    value.Availability,
                    value.Document == null ? null : $"{value.Document.Kind}:{value.Document.ServerRelativeUrl}",
                    value.Values.Count,
                    value.Attachments.Count,
                    Join(value.Diagnostics))));
            foreach (var item in source.Items.OrderBy(value => value.SourceItemId))
            {
                writer.Heading(5, "Item " + item.SourceItemId);
                if (item.Document != null)
                {
                    writer.Table(null, new[] { "Document property", "Value", "How to read it" }, new[]
                    {
                        Row("kind", item.Document.Kind, "Folder creates hierarchy; File materializes exact current bytes."),
                        Row("name", item.Document.Name, "Leaf name."),
                        Row("serverRelativeUrl", item.Document.ServerRelativeUrl, "Mapped relative to the source and target List roots."),
                        Row("length", item.Document.Length, "Source byte count for files."),
                        Row("majorVersion / minorVersion", $"{item.Document.MajorVersion} / {item.Document.MinorVersion}", "Captured evidence; version history is not replayed."),
                        Row("content", FormatArtifact(item.Document.Content), "Exact bytes may be inline Base64 or content-addressed in the artifact store."),
                        Row("content.representationKind", item.Document.Content?.RepresentationKind, "OrdinaryFilePayload is byte-stable evidence; InformationRightsManagedEnvelope is the exact protected response envelope and requires a separate replay decision."),
                        Row("content.logicalContentIdentity.quickXorHash", item.Document.Content?.LogicalContentIdentity?.QuickXorHash, "Stable source-content change identity exposed by SharePoint metadata; it does not replace artifact integrity SHA-256."),
                        Row("content.logicalContentIdentity.contentTag", item.Document.Content?.LogicalContentIdentity?.ContentTag, "Captured source content tag used with version, length, and QuickXorHash when comparing protected source reads."),
                        Row("content.logicalContentIdentity.evidenceSource", item.Document.Content?.LogicalContentIdentity?.EvidenceSource, "Where the stable logical-content identity was observed."),
                        Row("informationProtection.labelId", item.Document.InformationProtection?.LabelId, "Exact source Microsoft Information Protection label identity; it is not remapped by display name."),
                        Row("informationProtection.assignmentMethod", item.Document.InformationProtection?.AssignmentMethod, "Exact source assignment-method code retained from item metadata."),
                        Row("informationProtection.hasUserDefinedProtection", item.Document.InformationProtection?.HasUserDefinedProtection, "Exact source protection flag retained independently from the library IRM setting."),
                        Row("informationProtection.ownerEmail", item.Document.InformationProtection?.OwnerEmail, "Source label-owner evidence; not interpreted as a target principal mapping."),
                        Row("informationProtection.labelHash", item.Document.InformationProtection?.LabelHash, "Source label hash retained for change comparison."),
                        Row("informationProtection.promotionCtagVersion", item.Document.InformationProtection?.PromotionCtagVersion, "Source label-promotion version evidence."),
                        Row("informationProtection.decryptSkipReason", item.Document.InformationProtection?.DecryptSkipReason, "Source parser/decryption handling code extracted from MetaInfo."),
                        Row("archivedContentEvidence", FormatArchivedContentEvidence(item.Document.Content), "Literal HTTP 423 locked/contentArchived evidence means source reactivation and a fresh capture are required; it is not an authorization block.")
                    });
                }
                writer.Table(null, new[] { "Internal name", "Kind", "Typed value", "Raw runtime type", "Raw text", "Raw JSON", "Availability", "Diagnostics" },
                    item.Values.OrderBy(value => value.InternalName, StringComparer.OrdinalIgnoreCase).Select(value => Row(
                        value.InternalName,
                        value.Kind,
                        SummarizeListItemValue(value),
                        value.RawType,
                        Summarize(value.RawValue),
                        Summarize(value.RawValueJson),
                        value.Availability,
                        Join(value.Diagnostics))));
                writer.Table(null, new[] { "Attachment file", "Source path", "Content", "Availability / diagnostics" },
                    item.Attachments.OrderBy(value => value.FileName, StringComparer.OrdinalIgnoreCase).Select(value => Row(
                        value.FileName,
                        value.ServerRelativeUrl,
                        FormatArtifact(value.Content),
                        value.Content == null ? null : $"{value.Content.Availability}; archived={FormatArchivedContentEvidence(value.Content)}; {Join(value.Content.Diagnostics)}")));
            }
        }

        private static string FormatArchivedContentEvidence(ListBinaryArtifactSnapshot binary)
        {
            return binary?.ArchivedContentEvidence == null
                ? null
                : string.Join("; ", binary.ArchivedContentEvidence.Select(value =>
                    $"{value.Operation}: HTTP {value.HttpStatusCode} {value.ErrorCode}/{value.InnerErrorCode}; {value.RequestUri}; {value.ObservedAtUtc:O}; sha256={value.EvidenceSha256}"));
        }
    }
}
