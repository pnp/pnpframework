using PnP.Framework.Migration.Diagnostics;
using PnP.Framework.Migration.Schema.ContentTypes;
using PnP.Framework.Migration.Schema.Fields;
using System;
using System.Collections.Generic;
using System.Linq;

namespace PnP.Framework.Migration.Pages.Publishing.Reporting
{
    internal static class PublishingPageSchemaReportSection
    {
        public static void AppendSnapshot(MarkdownReportWriter writer, ContentTypeSchemaSnapshot schema)
        {
            if (schema == null)
            {
                writer.Table("Associated Page Layout content type schema",
                    new[] { "Property", "Value", "How to read it" },
                    new[] { Row("snapshot.layout.associatedContentTypeSchema", null, "No schema closure was captured; a custom Page Layout plan will block.") });
                return;
            }

            writer.Table("Associated Page Layout content type schema", new[] { "Property", "Value", "How to read it" }, new[]
            {
                Row("schemaVersion", schema.SchemaVersion, "Version of the reusable content-type schema evidence contract."),
                Row("evidenceState", schema.EvidenceState, "Readable means the selected field closure is complete; Partial retains exact identity, captured links/fields, and diagnostics without claiming a complete schema."),
                Row("availability", schema.Availability, "Captured can be materialized; Partial is eligible only for explicit exact target-runtime reuse when every captured field is target-runtime owned."),
                Row("sourceWebUrl", schema.SourceWebUrl, "Web from which the site content type schema was read."),
                Row("contentTypeId", schema.ContentTypeId, "Exact source content type ID to preserve when migration owns the type."),
                Row("name", schema.Name, "Source content type name."),
                Row("description", schema.Description, "Source content type description."),
                Row("group", schema.Group, "Source content type group."),
                Row("readOnly", schema.ReadOnly, "Read-only schema cannot be recreated as migration-owned without an explicit capability."),
                Row("sealed", schema.Sealed, "Sealed schema cannot be recreated as migration-owned without an explicit capability."),
                Row("hidden", schema.Hidden, "Source content type visibility."),
                Row("parentContentTypeId", schema.ParentContentTypeId, "Exact target parent must already be available."),
                Row("parentContentTypeName", schema.ParentContentTypeName, "Human-readable parent evidence."),
                Row("sources", PublishingPageArtifactReportFormatter.Sources(schema.Sources), "Lineage of the schema evidence."),
                Row("diagnostics", Join(schema.Diagnostics), "Capture conflicts or failures retained with the schema.")
            });

            writer.Table("Associated content type required field links",
                new[] { "Field ID", "Name", "Required", "Hidden", "Role", "Interpretation" },
                schema.RequiredFieldLinks.Select(item => Row(
                    item.FieldId,
                    item.Name,
                    item.Required,
                    item.Hidden,
                    item.Role,
                    item.Role == FieldSchemaRole.DirectBinding
                        ? "The associated content type directly exposes this captured field link."
                        : item.Role == FieldSchemaRole.InheritedFromParent
                            ? "The captured field link is inherited from the associated content type parent."
                            : "This field is required by another field in the captured closure.")));

            writer.Table("Associated content type required field-schema closure",
                new[] { "Field ID / internal name", "Title / type / group", "Flags", "Role / ownership", "Schema XML", "Schema SHA-256", "Portable SHA-256", "Taxonomy binding", "Sources", "Diagnostics" },
                schema.RequiredFieldClosure.Select(item => Row(
                    $"{item.Id:D} / {Format(item.InternalName)}",
                    $"{Format(item.Title)} / {Format(item.TypeAsString)} / {Format(item.Group)}",
                    $"required={item.Required}; hidden={item.Hidden}; readOnly={item.ReadOnly}; sealed={item.Sealed}",
                    $"role={item.Role}; ownership={item.Ownership}",
                    Summarize(item.SchemaXml),
                    item.SchemaXmlSha256,
                    item.PortableSchemaSha256,
                    PublishingPageArtifactReportFormatter.Taxonomy(item.Taxonomy),
                    PublishingPageArtifactReportFormatter.Sources(item.Sources),
                    Join(item.Diagnostics))));
        }

        public static void AppendPlan(MarkdownReportWriter writer, ContentTypeMaterializationPlan schema)
        {
            if (schema == null)
            {
                writer.Table("Page Layout content type materialization plan",
                    new[] { "Property", "Value", "How to read it" },
                    new[] { Row("plan.layoutMaterialization.contentTypeSchema", null, "Expected for stock-layout reuse or an explicitly blocked custom layout.") });
                return;
            }

            writer.Table("Page Layout content type materialization plan", new[] { "Property", "Value", "How to read it" }, new[]
            {
                Row("disposition", schema.Disposition, "CreateOwned permits create-or-exact-reuse from complete source schema; ReuseOwned requires an exact existing target-runtime content type and forbids schema writes; local Block records a capability gap and is projected as final ingredient Defer."),
                Row("sourceWebUrl", schema.SourceWebUrl, "Source schema provenance."),
                Row("contentTypeId", schema.ContentTypeId, "Exact ID to create or reuse."),
                Row("name", schema.Name, "Expected exact content type name."),
                Row("description", schema.Description, "Expected description."),
                Row("group", schema.Group, "Expected content type group."),
                Row("readOnly / sealed / hidden", $"{schema.ReadOnly} / {schema.Sealed} / {schema.Hidden}", "Expected exact content type flags."),
                Row("parentContentTypeId", schema.ParentContentTypeId, "Parent must already exist at the target."),
                Row("parentContentTypeName", schema.ParentContentTypeName, "Parent review label."),
                Row("reason", schema.Reason, "Human-readable schema policy decision.")
            });

            writer.Table("Planned content type field links",
                new[] { "Field ID", "Name", "Required", "Hidden", "Role" },
                schema.RequiredFieldLinks.Select(item => Row(item.FieldId, item.Name, item.Required, item.Hidden, item.Role)));
            writer.Table("Planned Page Layout field schemas",
                new[] { "Field ID / name", "Title / type / group", "Flags", "Role / ownership / disposition", "Source portable SHA-256", "Target schema XML", "Target portable SHA-256", "Source taxonomy", "Target taxonomy", "Hidden text field", "Reason" },
                schema.Fields.Select(item => Row(
                    $"{item.FieldId:D} / {Format(item.InternalName)}",
                    $"{Format(item.Title)} / {Format(item.TypeAsString)} / {Format(item.Group)}",
                    $"required={item.Required}; hidden={item.Hidden}",
                    $"role={item.Role}; ownership={item.Ownership}; disposition={item.Disposition}",
                    item.SourcePortableSchemaSha256,
                    Summarize(item.TargetSchemaXml),
                    item.TargetPortableSchemaSha256,
                    $"store={Format(item.SourceTermStoreId)}; set={Format(item.SourceTermSetId)}",
                    $"store={Format(item.TargetTermStoreId)}; set={Format(item.TargetTermSetId)}",
                    item.HiddenTextFieldId,
                    item.Reason)));
        }

        public static void AppendProbe(MarkdownReportWriter writer, ContentTypeTargetProbe probe)
        {
            if (probe == null)
            {
                writer.Table("Page Layout content type target probe",
                    new[] { "Property", "Value", "How to read it" },
                    new[] { Row("contentTypeSchema", null, "Expected for stock-layout reuse or an explicitly blocked custom layout.") });
                return;
            }

            writer.Table("Page Layout content type target probe", new[] { "Property", "Value", "How to read it" }, new[]
            {
                Row("contentTypeId", probe.ContentTypeId, "Exact planned ID."),
                Row("parentContentTypeAvailable", probe.ParentContentTypeAvailable, "Whether the required parent exists."),
                Row("resolvedParentContentTypeId", probe.ResolvedParentContentTypeId, "Parent actually resolved at the target."),
                Row("contentTypeExists", probe.ContentTypeExists, "Whether the exact planned ID already exists."),
                Row("existingName", probe.ExistingName, "Existing exact-ID metadata."),
                Row("existingDescription", probe.ExistingDescription, "Existing exact-ID metadata."),
                Row("existingGroup", probe.ExistingGroup, "Existing exact-ID metadata."),
                Row("existingReadOnly", probe.ExistingReadOnly, "Existing exact-ID metadata."),
                Row("existingSealed", probe.ExistingSealed, "Existing exact-ID metadata."),
                Row("existingHidden", probe.ExistingHidden, "Existing exact-ID metadata."),
                Row("existingParentContentTypeId", probe.ExistingParentContentTypeId, "Existing exact-ID parent."),
                Row("sameNameDifferentIds", Join(probe.SameNameDifferentIds), "Name collisions do not satisfy exact-ID reuse."),
                Row("canManageContentTypes", probe.CanManageContentTypes, "Effective target permission for schema creation."),
                Row("availability", probe.Availability, "Captured is required for admission."),
                Row("diagnostics", Join(probe.Diagnostics), "Target schema inspection diagnostics.")
            });
            writer.Table("Target parent content type field links",
                new[] { "Field ID", "Name", "Required", "Hidden" },
                probe.ParentFieldLinks.Select(item => Row(item.FieldId, item.Name, item.Required, item.Hidden)));
            writer.Table("Existing exact-ID content type field links",
                new[] { "Field ID", "Name", "Required", "Hidden" },
                probe.ExistingFieldLinks.Select(item => Row(item.FieldId, item.Name, item.Required, item.Hidden)));
            writer.Table("Target Page Layout field probes",
                new[] { "Field ID", "Exists", "Internal name", "Title", "Type", "Portable SHA-256" },
                probe.Fields.Select(item => Row(
                    item.FieldId,
                    item.Exists,
                    item.InternalName,
                    item.Title,
                    item.TypeAsString,
                    item.PortableSchemaSha256)));
        }

        public static void AppendAdmissionIssues(
            MarkdownReportWriter writer,
            ContentTypeTargetAdmission admission)
        {
            AppendIssues(writer, "Page Layout content type admission issues", admission?.Issues);
        }

        private static void AppendIssues(
            MarkdownReportWriter writer,
            string heading,
            IEnumerable<MigrationIssue> issues)
        {
            writer.Table(heading,
                new[] { "Code", "Severity", "Subject", "Ingredient", "Message", "Source identity", "Target identity" },
                (issues ?? Array.Empty<MigrationIssue>()).Select(item => Row(
                    item.Code,
                    item.Severity,
                    item.Subject,
                    item.Ingredient,
                    item.Message,
                    item.SourceIdentity,
                    item.TargetIdentity)));
        }

        private static string[] Row(params object[] values) => PublishingPageReportValueFormatter.Row(values);

        private static string Format(object value) => PublishingPageReportValueFormatter.Format(value);

        private static string Join(IEnumerable<string> values) => PublishingPageReportValueFormatter.Join(values);

        private static string Summarize(string value) => PublishingPageReportValueFormatter.SummarizePayload(value);
    }
}
