using PnP.Framework.Migration.Pages.Capture;
using PnP.Framework.Migration.Pages.Publishing.Packaging;
using PnP.Framework.Migration.Pages.Publishing.Capture;
using PnP.Framework.Migration.Pages.Publishing.Planning;
using PnP.Framework.Migration.Pages.Planning;
using System;
using System.Linq;

namespace PnP.Framework.Migration.Pages.Publishing.Reporting
{
    public static class PublishingPageMigrationReportBuilder
    {
        public static string Build(PublishingPageMigrationPackage package)
        {
            PublishingPagePackageValidator.ValidateMigration(package);
            var snapshot = package.Snapshot;
            var plan = package.Plan;
            var report = package.Report ?? new PublishingPageMigrationReport();
            var writer = new MarkdownReportWriter();
            writer.Heading(1, $"{DisplayProfile(snapshot.SourceProfile)} migration report");
            writer.Paragraph(report.Summary ?? "This report describes the sealed source snapshot and target-specific migration plan.");
            writer.Paragraph("The JSON package is authoritative. This Markdown is a complete review view: large HTML, XML, JSON, and Base64 values are represented by length, SHA-256, and a preview while their full values remain in the package.");

            writer.Table("Package envelope", new[] { "JSON path", "Value", "How to read it" }, new[]
            {
                Row("schemaVersion", package.SchemaVersion, "Version of the planned/importable package contract."),
                Row("plannedAtUtc", package.PlannedAtUtc, "UTC time at which the target-specific plan was sealed."),
                Row("exportSchemaVersion", package.ExportSchemaVersion, "Version of the embedded source-only export contract."),
                Row("exportedAtUtc", package.ExportedAtUtc, "UTC time at which the source snapshot was captured."),
                Row("state", package.State, "ApprovalReady can be imported after digest approval; Blocked cannot."),
                Row("snapshotDigest", package.SnapshotDigest, "SHA-256 over the complete source snapshot."),
                Row("planDigest", package.PlanDigest, "SHA-256 over all target decisions; this is the approval token."),
                Row("snapshot.sourceProfile", snapshot.SourceProfile, "Selects the source-specific classifier and target adapter."),
                Row("plan.isExecutable", plan.IsExecutable, "True only when the blocker list is empty."),
                Row("report.summary", report.Summary, "Human-readable package status summary.")
            });

            writer.Table("Source page identity", new[] { "Property", "Value", "How to read it" }, new[]
            {
                Row("webUrl", snapshot.Source.WebUrl, "Absolute source web URL."),
                Row("webServerRelativeUrl", snapshot.Source.WebServerRelativeUrl, "Server-relative source web boundary used for URL mapping."),
                Row("pageServerRelativeUrl", snapshot.Source.PageServerRelativeUrl, "Exact source page path."),
                Row("listItemId", snapshot.Source.ListItemId, "Source Pages-library item ID; evidence only across sites."),
                Row("fileUniqueId", snapshot.Source.FileUniqueId, "Source file GUID used as identity evidence."),
                Row("contentTypeId", snapshot.Source.ContentTypeId, "Source content type used by the selected migration profile."),
                Row("contentTypeName", snapshot.Source.ContentTypeName, "Human-readable source content type."),
                Row("versionLabel", snapshot.Source.VersionLabel, "Version captured by the export."),
                Row("length", snapshot.Source.Length, "Source ASPX file length in bytes."),
                Row("modifiedUtc", snapshot.Source.ModifiedUtc, "Source file modified time at capture."),
                Row("title", snapshot.Source.Title, "Title assigned during target page creation."),
                Row("layout.url", snapshot.Layout.Url, "Source publishing layout URL."),
                Row("layout.description", snapshot.Layout.Description, "Source layout description, if present.")
            });

            writer.Table("Capture policy and source fence", new[] { "Property", "Value", "How to read it" }, new[]
            {
                Row("capturePolicy.sourcePageServerRelativeUrl", snapshot.CapturePolicy.SourcePageServerRelativeUrl, "Normalized source page requested at export."),
                Row("capturePolicy.includeWebParts", snapshot.CapturePolicy.IncludeWebParts, "Whether shared Web Parts were inventoried and exported."),
                Row("capturePolicy.maximumDependencyBytes", snapshot.CapturePolicy.MaximumDependencyBytes, "Maximum bytes captured for each restorable dependency."),
                Row("sourceFence.fileUniqueId", snapshot.SourceFence.FileUniqueId, "Identity checked before and after export."),
                Row("sourceFence.versionLabel", snapshot.SourceFence.VersionLabel, "Version checked before and after export."),
                Row("sourceFence.length", snapshot.SourceFence.Length, "File length checked before and after export."),
                Row("sourceFence.modifiedUtc", snapshot.SourceFence.ModifiedUtc, "Modified time checked before and after export.")
            });

            writer.Table("Lifecycle decision", new[] { "Property", "Value", "How to read it" }, new[]
            {
                Row("snapshot.lifecycle.checkOutType", snapshot.Lifecycle.CheckOutType, "Source checkout state captured as evidence."),
                Row("snapshot.lifecycle.level", snapshot.Lifecycle.Level, "Only an unconflicted Published value maps to Published; every other state maps to Draft."),
                Row("snapshot.lifecycle.moderationStatus", snapshot.Lifecycle.ModerationStatus, "Source moderation status value, if available."),
                Row("snapshot.lifecycle.createdUtc", snapshot.Lifecycle.CreatedUtc, "Source creation time; not replayed as a system field."),
                Row("snapshot.lifecycle.modifiedUtc", snapshot.Lifecycle.ModifiedUtc, "Source modified time; not replayed as a system field."),
                Row("plan.targetLifecycle", plan.TargetLifecycle, "Derived lifecycle to enforce at the target."),
                Row("plan.lifecycleReason", plan.LifecycleReason, "Human-readable derivation rule; there is no publish Boolean input.")
            });

            writer.Table("Publishing content", new[] { "Property", "Value", "How to read it" }, new[]
            {
                Row("snapshot.publishingPageContent", Summarize(snapshot.PublishingPageContent), "Full HTML remains in JSON; this row gives length, digest, and preview."),
                Row("snapshot.publishingPageContentSha256", snapshot.PublishingPageContentSha256, "Digest of source HTML before URL replacements."),
                Row("plan.expectedPublishingPageContentSha256", plan.ExpectedPublishingPageContentSha256, "Digest expected after approved replacements and before SharePoint serialization.")
            });

            AppendFields(writer, snapshot, plan);
            AppendWebParts(writer, snapshot);
            AppendDependencies(writer, snapshot, plan);
            AppendSecurity(writer, snapshot);
            AppendPlan(writer, plan);

            writer.Table("Approved text replacements", new[] { "Source", "Target", "Reason" },
                plan.Replacements.Select(item => Row(item.Source, item.Target, item.Reason)));
            writer.List("Storage assertions", plan.StorageAssertions);
            writer.List("Browser acceptance assertions", plan.BrowserAssertions);
            writer.List("Plan blockers", plan.Blockers);
            writer.List("Plan warnings", plan.Warnings);
            writer.List("Snapshot blockers", snapshot.Blockers);
            writer.List("Snapshot warnings", snapshot.Warnings);
            writer.List("Report blockers", report.Blockers);
            writer.List("Report warnings", report.Warnings);
            writer.List("Captured ingredients", report.CapturedIngredients);
            writer.Heading(2, "Field-action legend");
            writer.Paragraph("`Apply` is written by Import. `AlreadyHandled` is handled by page creation/content/layout logic. `EvidenceOnly` remains available for future recovery. `RequiresMapping` needs an explicit cross-site identity mapping. `Skip*`, `Target*`, and `CaptureUnavailable` are retained but not written. `Block` makes the plan non-executable.");
            return writer.ToString();
        }

        private static void AppendFields(
            MarkdownReportWriter writer,
            PublishingPageCaptureBundle snapshot,
            PublishingPageMigrationPlan plan)
        {
            var actionByField = plan.FieldActions.ToDictionary(item => item.SourceInternalName, StringComparer.OrdinalIgnoreCase);
            writer.Table($"Complete list-item field inventory ({snapshot.Fields.Count})",
                new[] { "#", "Internal name", "Title / ID", "Source type", "Capture", "Value", "Flags", "Plan action", "Will apply", "Target", "Reason" },
                snapshot.Fields.Select((field, index) =>
                {
                    actionByField.TryGetValue(field.InternalName, out var action);
                    return Row(
                        index + 1,
                        field.InternalName,
                        $"{Format(field.Title)} / {field.Id:D}",
                        field.TypeAsString,
                        $"{field.CaptureStatus}; hasValue={field.HasValue}; rawType={Format(field.RawType)}",
                        PublishingPageReportValueFormatter.SummarizeFieldValue(field),
                        $"readOnly={field.ReadOnly}; hidden={field.Hidden}; required={field.Required}",
                        action?.Disposition,
                        action?.WillApply,
                        action == null ? null : $"{Format(action.TargetInternalName)} ({Format(action.TargetTypeAsString)})",
                        action?.Reason);
                }));

            writer.Heading(2, "Per-field recovery evidence");
            foreach (var field in snapshot.Fields)
            {
                writer.Heading(3, PublishingPageReportValueFormatter.EscapeHeading(field.InternalName));
                writer.Table(null, new[] { "Property", "Value" }, new[]
                {
                    Row("id", field.Id),
                    Row("internalName", field.InternalName),
                    Row("title", field.Title),
                    Row("typeAsString", field.TypeAsString),
                    Row("schemaXml", Summarize(field.SchemaXml)),
                    Row("readOnly", field.ReadOnly),
                    Row("hidden", field.Hidden),
                    Row("required", field.Required),
                    Row("hasValue", field.HasValue),
                    Row("kind", field.Kind),
                    Row("value", field.Value),
                    Row("stringValues", Join(field.StringValues)),
                    Row("urlValue", field.UrlValue == null ? null : $"url={Format(field.UrlValue.Url)}; description={Format(field.UrlValue.Description)}"),
                    Row("lookupValues", Join(field.LookupValues.Select(value => $"id={value.LookupId}; value={Format(value.LookupValue)}"))),
                    Row("taxonomyValues", Join(field.TaxonomyValues.Select(value => $"label={Format(value.Label)}; termGuid={Format(value.TermGuid)}; wssId={value.WssId}"))),
                    Row("binaryBase64", Summarize(field.BinaryBase64)),
                    Row("rawType", field.RawType),
                    Row("rawValue", Summarize(field.RawValue)),
                    Row("rawValueJson", Summarize(field.RawValueJson)),
                    Row("captureStatus", field.CaptureStatus),
                    Row("diagnostics", Join(field.Diagnostics))
                });
            }
        }

        private static void AppendWebParts(MarkdownReportWriter writer, PublishingPageCaptureBundle snapshot)
        {
            writer.Table($"Shared Web Parts ({snapshot.WebParts.Count})",
                new[] { "ID", "Title", "Zone", "Index", "Hidden", "Export SHA-256", "Export XML" },
                snapshot.WebParts.Select(item => Row(
                    item.Id,
                    item.Title,
                    item.ZoneId,
                    item.ZoneIndex,
                    item.Hidden,
                    item.ExportSha256,
                    Summarize(item.ExportXml))));
        }

        private static void AppendDependencies(
            MarkdownReportWriter writer,
            PublishingPageCaptureBundle snapshot,
            PublishingPageMigrationPlan plan)
        {
            var actionByDependency = plan.DependencyActions.ToDictionary(item => item.SnapshotDependencyId, StringComparer.Ordinal);
            writer.Table($"Dependencies ({snapshot.Dependencies.Count})",
                new[] { "ID", "Consumer / kind", "Original", "Source absolute / path", "Capture", "Payload", "Plan action", "Target absolute / path", "Diagnostics" },
                snapshot.Dependencies.Select(item =>
                {
                    actionByDependency.TryGetValue(item.Id, out var action);
                    return Row(
                        item.Id,
                        $"{item.Consumer} / {item.Kind}; renderable={item.IsRenderableResource}",
                        item.OriginalValue,
                        $"absolute={Format(item.SourceAbsoluteUrl)}; path={Format(item.SourceServerRelativeUrl)}",
                        item.CaptureStatus,
                        $"bytes={item.ContentLength}; sha256={Format(item.ContentSha256)}; base64={Summarize(item.ContentBase64)}",
                        action?.Disposition,
                        action == null ? null : $"absolute={Format(action.TargetAbsoluteUrl)}; path={Format(action.TargetServerRelativeUrl)}",
                        Join(item.Diagnostics.Concat(action?.Diagnostics ?? Array.Empty<string>())));
                }));
        }

        private static void AppendSecurity(MarkdownReportWriter writer, PublishingPageCaptureBundle snapshot)
        {
            writer.Table("Security snapshot", new[] { "Property", "Value", "How to read it" }, new[]
            {
                Row("hasUniqueRoleAssignments", snapshot.Security.HasUniqueRoleAssignments, "Unique permissions are captured; the current importer does not replay them."),
                Row("roleAssignmentCount", snapshot.Security.RoleAssignments.Count, "Number of captured source role assignments.")
            });
            writer.Table("Source role assignments", new[] { "Principal login", "Principal title", "Role definitions" },
                snapshot.Security.RoleAssignments.Select(item => Row(
                    item.PrincipalLoginName,
                    item.PrincipalTitle,
                    Join(item.RoleDefinitionNames))));
        }

        private static void AppendPlan(MarkdownReportWriter writer, PublishingPageMigrationPlan plan)
        {
            writer.Table("Plan and target probe", new[] { "Property", "Value", "How to read it" }, new[]
            {
                Row("plan.operation", plan.Operation, "CreatePage is executable now; ApplyDeferredFields reserves future snapshot-based recovery."),
                Row("plan.sourceSnapshotDigest", plan.SourceSnapshotDigest, "Binds this plan to the exact source snapshot."),
                Row("plan.sourceWebUrl", plan.SourceWebUrl, "Source web used for mapping."),
                Row("plan.sourcePageServerRelativeUrl", plan.SourcePageServerRelativeUrl, "Source page represented by the snapshot."),
                Row("plan.targetWebUrl", plan.TargetWebUrl, "Target connection must point here during import."),
                Row("plan.targetWebServerRelativeUrl", plan.TargetWebServerRelativeUrl, "Target web boundary for paths."),
                Row("plan.targetPageServerRelativeUrl", plan.TargetPageServerRelativeUrl, "Exact target page to create."),
                Row("plan.pageLayoutName", plan.PageLayoutName, "Publishing layout name passed to page creation."),
                Row("plan.createOnly", plan.CreateOnly, "Existing pages and dependency files are blockers."),
                Row("planningPolicy.targetPageServerRelativeUrl", plan.PlanningPolicy.TargetPageServerRelativeUrl, "Normalized requested target page."),
                Row("planningPolicy.requireInheritedPermissions", plan.PlanningPolicy.RequireInheritedPermissions, "Blocks source pages with unique permissions when true."),
                Row("planningPolicy.blockOnManagedMetadata", plan.PlanningPolicy.BlockOnManagedMetadata, "Blocks non-empty taxonomy values without a reviewed term mapping when true."),
                Row("planningPolicy.allowExternalResourceReferences", plan.PlanningPolicy.AllowExternalResourceReferences, "Allows external authored resource references to remain external when true."),
                Row("planningPolicy.createOnly", plan.PlanningPolicy.CreateOnly, "Disallows replacing target files when true."),
                Row("targetProbe.webUrl", plan.TargetProbe.WebUrl, "Web actually probed while planning."),
                Row("targetProbe.webServerRelativeUrl", plan.TargetProbe.WebServerRelativeUrl, "Target web path boundary."),
                Row("targetProbe.webTemplate", plan.TargetProbe.WebTemplate, "Target web template identifier."),
                Row("targetProbe.webConfiguration", plan.TargetProbe.WebConfiguration, "Target web configuration number."),
                Row("targetProbe.pagesLibraryServerRelativeUrl", plan.TargetProbe.PagesLibraryServerRelativeUrl, "Publishing Pages library root."),
                Row("targetProbe.pagesLibraryBaseTemplate", plan.TargetProbe.PagesLibraryBaseTemplate, "Expected publishing template value is 850."),
                Row("targetProbe.enableVersioning", plan.TargetProbe.EnableVersioning, "Required for deterministic lifecycle handling."),
                Row("targetProbe.enableMinorVersions", plan.TargetProbe.EnableMinorVersions, "Required when the derived target lifecycle is Draft."),
                Row("targetProbe.enableModeration", plan.TargetProbe.EnableModeration, "Controls whether a published page also needs approval."),
                Row("targetProbe.forceCheckout", plan.TargetProbe.ForceCheckout, "Controls whether the importer must check out before updates."),
                Row("targetProbe.draftVersionVisibility", plan.TargetProbe.DraftVersionVisibility, "Who can see draft/minor versions."),
                Row("targetProbe.pageContentTypeId", plan.TargetProbe.PageContentTypeId, "Profile-specific publishing content type found in the target library."),
                Row("targetProbe.pageLayoutUrl", plan.TargetProbe.PageLayoutUrl, "Profile-specific layout found in the target site collection."),
                Row("targetProbe.pageLayoutExists", plan.TargetProbe.PageLayoutExists, "False is a blocker."),
                Row("targetProbe.targetPageExists", plan.TargetProbe.TargetPageExists, "True is a blocker for CreatePage."),
                Row("targetProbe.existingDependencyPaths", Join(plan.TargetProbe.ExistingDependencyPaths), "Create-only dependency collisions.")
            });
        }

        private static string DisplayProfile(string value)
        {
            return string.Equals(value, "EnterpriseWiki", StringComparison.Ordinal)
                ? "Enterprise Wiki"
                : Format(value);
        }

        private static string[] Row(params object[] values) => PublishingPageReportValueFormatter.Row(values);

        private static string Format(object value) => PublishingPageReportValueFormatter.Format(value);

        private static string Join(System.Collections.Generic.IEnumerable<string> values) => PublishingPageReportValueFormatter.Join(values);

        private static string Summarize(string value) => PublishingPageReportValueFormatter.SummarizePayload(value);
    }
}
