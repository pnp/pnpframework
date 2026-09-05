using PnP.Framework.Migration.Diagnostics;
using PnP.Framework.Migration.Pages.Publishing.Layouts;
using PnP.Framework.Migration.Pages.Publishing.Planning;
using System;
using System.Collections.Generic;
using System.Linq;

namespace PnP.Framework.Migration.Pages.Publishing.Reporting
{
    internal static class PublishingPageLayoutPlanReportSection
    {
        public static void Append(MarkdownReportWriter writer, PublishingPageMigrationPlan pagePlan)
        {
            var plan = pagePlan.LayoutMaterialization;
            writer.Table("Page Layout materialization plan", new[] { "JSON path", "Value", "How to read it" }, new[]
            {
                Row("plan.layoutMaterialization.disposition", plan.Disposition, "ReuseTargetStock uses reviewed target runtime stock; CreateOwned creates a digest-owned layout; ReuseOwned is a fresh admission result; local Block becomes final ingredient Defer and remains queued for mitigation unless literal HTTP 401/403 evidence exists."),
                Row("plan.layoutMaterialization.sourceUrl", plan.SourceUrl, "Absolute source layout URL sealed into the target decision."),
                Row("plan.layoutMaterialization.sourceServerRelativeUrl", plan.SourceServerRelativeUrl, "Source layout path."),
                Row("plan.layoutMaterialization.sourceFileName", plan.SourceFileName, "Original source file name."),
                Row("plan.layoutMaterialization.sourceBytes", PublishingPageArtifactReportFormatter.Artifact(plan.SourceBytes), "Exact captured source ASPX artifact."),
                Row("plan.layoutMaterialization.associatedContentTypeName", plan.AssociatedContentTypeName, "Associated publishing content type name."),
                Row("plan.layoutMaterialization.associatedContentTypeId", plan.AssociatedContentTypeId, "Source associated content type ID."),
                Row("plan.layoutMaterialization.targetFileName", plan.TargetFileName, "Stock file name or deterministic pnp-{source-stem}-{source-digest-prefix}.aspx."),
                Row("plan.layoutMaterialization.targetPageLayoutName", plan.TargetPageLayoutName, "Layout name passed to page creation; must equal plan.pageLayoutName."),
                Row("plan.layoutMaterialization.targetServerRelativeUrl", plan.TargetServerRelativeUrl, "Exact target master-page-gallery path."),
                Row("plan.layoutMaterialization.targetBytes", PublishingPageArtifactReportFormatter.Artifact(plan.TargetBytes), "Expected bytes after approved resource-reference rewrites."),
                Row("plan.layoutMaterialization.requiredFieldBindings", Join(plan.RequiredFieldBindings), "Minimal field names bound by parsed layout controls."),
                Row("plan.layoutMaterialization.reason", plan.Reason, "Human-readable policy decision.")
            });

            writer.Table("Planned Page Layout registrations",
                new[] { "Tag prefix", "Namespace", "Assembly" },
                plan.RequiredRegistrations.Select(item => Row(item.TagPrefix, item.Namespace, item.Assembly)));
            writer.Table("Planned Page Layout zones",
                new[] { "ID", "Title" },
                plan.Zones.Select(item => Row(item.Id, item.Title)));
            writer.Table("Planned Page Layout resource references",
                new[] { "Attribute", "Authored value" },
                plan.ResourceReferences.Select(item => Row(item.Attribute, item.Value)));
            writer.Table("Page Layout resource materialization actions",
                new[] { "Source reference", "Source URL", "Source evidence", "Disposition", "Source artifact", "Inline bytes", "Target path", "Target reference", "Reason" },
                plan.ResourceMaterializations.Select(item => Row(
                    item.SourceReference,
                    item.SourceUrl,
                    item.SourceEvidenceState,
                    item.Disposition,
                    PublishingPageArtifactReportFormatter.Artifact(item.SourceArtifact),
                    Summarize(item.SourceContentBase64),
                    item.TargetServerRelativeUrl,
                    item.TargetReference,
                    item.Reason)));
            writer.Table("Page Layout resource rewrites",
                new[] { "Source authored reference", "Target reference", "Interpretation" },
                plan.ResourceRewrites.Select(item => Row(
                    item.SourceReference,
                    item.TargetReference,
                    "This exact ordinal string substitution is applied to the captured ASPX bytes before sealing targetBytes.")));

            PublishingPageSchemaReportSection.AppendPlan(writer, plan.ContentTypeSchema);
            AppendTargetProbe(writer, pagePlan.LayoutTargetProbe);
            AppendAdmission(writer, pagePlan.LayoutAdmission);
        }

        private static void AppendTargetProbe(MarkdownReportWriter writer, PublishingPageLayoutTargetProbe probe)
        {
            if (probe == null)
            {
                writer.Table("Page Layout target probe",
                    new[] { "Property", "Value", "How to read it" },
                    new[] { Row("plan.layoutTargetProbe", null, "No target layout evidence was sealed; an executable plan cannot omit it.") });
                return;
            }

            writer.Table("Page Layout target probe", new[] { "Property", "Value", "How to read it" }, new[]
            {
                Row("targetServerRelativeUrl", probe.TargetServerRelativeUrl, "Exact layout path inspected during planning."),
                Row("fileExists", probe.FileExists, "For target stock this must be true; for digest-owned layouts either absent or exact reuse is permitted."),
                Row("existingBytesSha256", probe.ExistingBytesSha256, "Fresh digest of an existing target layout, when present."),
                Row("existingAssociatedContentTypeName", probe.ExistingAssociatedContentTypeName, "Existing gallery association name."),
                Row("existingAssociatedContentTypeId", probe.ExistingAssociatedContentTypeId, "Existing gallery association ID."),
                Row("associatedContentTypeAvailable", probe.AssociatedContentTypeAvailable, "Whether an exact usable associated content type already exists."),
                Row("resolvedAssociatedContentTypeId", probe.ResolvedAssociatedContentTypeId, "Target content type ID selected for association."),
                Row("missingFieldBindings", Join(probe.MissingFieldBindings), "Required field bindings not presently exposed by the target root Web."),
                Row("canAddAndCustomizePages", probe.CanAddAndCustomizePages, "Effective permission required to create a custom ASPX layout."),
                Row("availability", probe.Availability, "Captured is required for admission."),
                Row("diagnostics", Join(probe.Diagnostics), "Target-inspection diagnostics.")
            });

            writer.Table("Page Layout target resource probes",
                new[] { "Target path", "Exists", "Existing SHA-256", "Can write", "Availability", "Diagnostics" },
                probe.Resources.Select(item => Row(
                    item.TargetServerRelativeUrl,
                    item.FileExists,
                    item.ExistingBytesSha256,
                    item.CanWrite,
                    item.Availability,
                    Join(item.Diagnostics))));
            PublishingPageSchemaReportSection.AppendProbe(writer, probe.ContentTypeSchema);
        }

        private static void AppendAdmission(MarkdownReportWriter writer, PublishingPageLayoutTargetAdmission admission)
        {
            writer.Table("Page Layout target admission", new[] { "Property", "Value", "How to read it" }, new[]
            {
                Row("isEligible", admission.IsEligible, "True only when every required layout, schema, resource, permission, and collision check is satisfied."),
                Row("disposition", admission.Disposition, "The target-observed action approved at planning time."),
                Row("contentTypeSchema.isEligible", admission.ContentTypeSchema?.IsEligible, "Custom layouts require a separately eligible schema admission."),
                Row("contentTypeSchema.disposition", admission.ContentTypeSchema?.Disposition, "CreateOwned, ReuseOwned, or a local capability Block for the associated schema; local Block is projected as final ingredient Defer."),
                Row("warnings", Join(admission.Warnings), "Non-blocking admission findings."),
                Row("contentTypeSchema.warnings", Join(admission.ContentTypeSchema?.Warnings), "Non-blocking schema findings.")
            });
            AppendIssues(writer, "Page Layout admission issues", admission.Issues);
            PublishingPageSchemaReportSection.AppendAdmissionIssues(writer, admission.ContentTypeSchema);
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

        private static string Join(IEnumerable<string> values) => PublishingPageReportValueFormatter.Join(values);

        private static string Summarize(string value) => PublishingPageReportValueFormatter.SummarizePayload(value);
    }
}
