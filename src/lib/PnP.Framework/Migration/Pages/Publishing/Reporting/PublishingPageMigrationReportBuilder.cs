using PnP.Framework.Migration.Packaging;
using PnP.Framework.Migration.Pages.Capture;
using PnP.Framework.Migration.Pages.Publishing.Capture;
using PnP.Framework.Migration.Pages.Publishing.Packaging;
using PnP.Framework.Migration.Pages.Publishing.Planning;
using PnP.Framework.Migration.Pages.Publishing.Reporting.Sections;
using PnP.Framework.Migration.Pages.Markup;
using PnP.Framework.Migration.Pages.Fields;
using PnP.Framework.Migration.Taxonomy;
using System;
using System.Linq;

namespace PnP.Framework.Migration.Pages.Publishing.Reporting
{
    public static class PublishingPageMigrationReportBuilder
    {
        public static string Build(PublishingPageMigrationPackage package)
        {
            return Build(package, null);
        }

        public static string Build(PublishingPageMigrationPackage package, IMigrationArtifactStore artifactStore)
        {
            PublishingPagePackageValidator.ValidateMigration(package, artifactStore);
            var snapshot = package.Snapshot;
            var plan = package.Plan;
            var report = package.Report ?? new PublishingPageMigrationReport();
            var writer = new MarkdownReportWriter();
            writer.Heading(1, $"{DisplayWorkflow(package.Selection.WorkflowId)} migration report");
            writer.Paragraph(report.Summary ?? "This report describes the sealed source snapshot and target-specific migration plan.");
            writer.Paragraph("The JSON package is authoritative. This Markdown is a complete review view: large HTML, XML, JSON, and Base64 values are represented by length, SHA-256, and a preview while their full values remain in the package.");

            writer.Table("Package envelope", new[] { "JSON path", "Value", "How to read it" }, new[]
            {
                Row("schemaVersion", package.SchemaVersion, "Version of the planned/importable package contract."),
                Row("plannedAtUtc", package.PlannedAtUtc, "UTC time at which the target-specific plan was sealed."),
                Row("exportSchemaVersion", package.ExportSchemaVersion, "Version of the embedded source-only export contract."),
                Row("exportedAtUtc", package.ExportedAtUtc, "UTC time at which the source snapshot was captured."),
                Row("state", package.State, "ApprovalReady can be imported after digest approval; MitigationPending is re-queued; AuthorizationBlocked requires literal HTTP 401/403 evidence; Invalid requires RCA."),
                Row("snapshotDigest", package.SnapshotDigest, "SHA-256 over the complete source snapshot."),
                Row("planDigest", package.PlanDigest, "SHA-256 over all target decisions; this is the approval token."),
                Row("selection.workflowId", package.Selection.WorkflowId, "Selects the orchestration and policy set; it does not redefine captured evidence."),
                Row("selectionDigest", package.SelectionDigest, "SHA-256 over workflow and cohort selection; editing classification invalidates the package."),
                Row("selection.validationCohort.cohortId", package.Selection.ValidationCohort.CohortId, "Names the validation population used by this workflow."),
                Row("selection.validationCohort.policyVersion", package.Selection.ValidationCohort.PolicyVersion, "Version of the cohort-membership policy."),
                Row("selection.validationCohort.disposition", package.Selection.ValidationCohort.Disposition, "Included means this page belongs to the reviewed EW-v1 validation cohort; capability is assessed separately per ingredient."),
                Row("selection.validationCohort.reasons", Join(package.Selection.ValidationCohort.Reasons), "Evidence-backed explanation of cohort membership."),
                Row("plan.migrationOutcome", plan.MigrationOutcome, "Aggregate result of ingredient actions and required dependency closure."),
                Row("plan.isExecutable", plan.IsExecutable, "True only when there are no workflow blockers and the ingredient outcome is executable."),
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
                Row("title", snapshot.Source.Title, "Title assigned during target page creation.")
            });

            AppendRuntimeAndClassification(writer, snapshot);
            AppendPageArtifact(writer, snapshot);

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

            PublishingPageLayoutSnapshotReportSection.Append(writer, snapshot.Layout);
            TopologyMigrationReportSection.Append(writer, snapshot, plan);
            AppendFields(writer, snapshot, plan);
            ListDependencyMigrationReportSection.Append(writer, snapshot, plan);
            ClassicWebPartMigrationReportSection.Append(writer, snapshot, plan);
            AppendDependencies(writer, snapshot, plan);
            AppendSecurity(writer, snapshot);
            AppendIngredientModel(writer, snapshot, plan);
            AppendPlan(writer, plan);

            writer.Table("Approved text replacements", new[] { "Source", "Target", "Reason" },
                plan.Replacements.Select(item => Row(item.Source, item.Target, item.Reason)));
            writer.List("Storage assertions", plan.StorageAssertions);
            writer.Table(
                "Runtime verification requirements",
                new[] { "ID", "Kind", "Required", "Description" },
                plan.RuntimeVerification.Requirements.Select(item => Row(item.Id, item.Kind, item.Required, item.Description)));
            writer.List("Plan blockers", plan.Blockers);
            writer.List("Plan warnings", plan.Warnings);
            writer.List("Snapshot blockers", snapshot.Blockers);
            writer.List("Snapshot warnings", snapshot.Warnings);
            writer.List("Report blockers", report.Blockers);
            writer.List("Report warnings", report.Warnings);
            writer.List("Captured ingredients", report.CapturedIngredients);
            writer.Heading(2, "Field-action legend");
            writer.Paragraph("`Apply` writes a supported scalar value. `ApplyTaxonomyRelationships` executes the separately reviewed relationship actions and never creates or substitutes a Term. A taxonomy relationship action of `RetainEvidenceOnly` preserves its sealed source proof without asserting target capability. `AlreadyHandled` is handled by page creation/content/layout logic. `EvidenceOnly` remains available for future recovery. `RequiresMapping` needs an explicit cross-site identity mapping. `Skip*`, `Target*`, and `CaptureUnavailable` are retained but not written. A typed field planner may use `Block` to record a current capability gap; final ingredient orchestration converts that finding to nonterminal `Defer` unless the same ingredient carries validated literal wire HTTP 401/403 evidence.");
            writer.Heading(2, "Ingredient-action legend");
            writer.Paragraph("`Preserve` retains the ingredient's semantics, `Transform` deliberately changes representation, `Substitute` lets the target runtime supply an equivalent, `Drop` records reviewed loss, `Delegate` keeps evidence for another workflow, and `Defer` keeps a known gap in the mitigation and re-planning queue. Final `Block` is reserved for digest-valid literal wire HTTP 401/403 evidence and stops only that affected ingredient branch. A retained ingredient may only drop a required dependency when its transform explicitly releases that dependency.");
            return writer.ToString();
        }

        private static void AppendRuntimeAndClassification(
            MarkdownReportWriter writer,
            PublishingPageCaptureBundle snapshot)
        {
            writer.Table("CLR runtime resolution", new[] { "Property", "Value", "How to read it" }, new[]
            {
                Row("schemaVersion", snapshot.Runtime.SchemaVersion, "Version of the runtime-resolution evidence contract."),
                Row("pageDeclaredType", snapshot.Runtime.PageDeclaredType, "CLR type declared by the source ASPX Page directive."),
                Row("layoutDeclaredType", snapshot.Runtime.LayoutDeclaredType, "CLR type declared by the Page Layout directive, when present."),
                Row("adapterId", snapshot.Runtime.AdapterId, "Executable adapter selected from CLR evidence; Content Type is only a fallback signal."),
                Row("detectionSource", snapshot.Runtime.DetectionSource, "Evidence source that won runtime resolution."),
                Row("resolutionState", snapshot.Runtime.ResolutionState, "Resolved is executable by a matching adapter; ambiguous or unknown evidence remains explicit."),
                Row("diagnostics", Join(snapshot.Runtime.Diagnostics), "Runtime resolution explanations and conflicts.")
            });

            writer.Table(
                $"Non-exclusive page profile signals ({snapshot.ProfileSignals.Count})",
                new[] { "Profile ID", "Signal kind", "Subject", "Evidence", "How to read it" },
                snapshot.ProfileSignals.Select(signal => Row(
                    signal.ProfileId,
                    signal.Kind,
                    signal.Subject,
                    signal.Evidence,
                    "Signals classify product ancestry or observed traits; multiple profiles may apply and none selects the runtime adapter.")));
        }

        private static void AppendPageArtifact(
            MarkdownReportWriter writer,
            PublishingPageCaptureBundle snapshot)
        {
            var artifact = snapshot.PageArtifact;
            writer.Table("Source ASPX artifact", new[] { "Property", "Value", "How to read it" }, new[]
            {
                Row("schemaVersion", artifact.SchemaVersion, "Version of the immutable page-artifact evidence contract."),
                Row("fileUniqueId", artifact.FileUniqueId, "Must match the source page identity."),
                Row("serverRelativeUrl", artifact.ServerRelativeUrl, "Must match the source page path."),
                Row("bytes", PublishingPageArtifactReportFormatter.Artifact(artifact.Bytes), "Reference to the exact source ASPX bytes."),
                Row("contentBase64", Summarize(artifact.ContentBase64), "Inline bytes when no external artifact store is used; the JSON retains the complete value."),
                Row("availability", artifact.Availability, "Captured means byte evidence passed digest validation."),
                Row("diagnostics", Join(artifact.Diagnostics), "Capture failures or partial-evidence explanations."),
                Row("pageDirective.inherits", artifact.PageDirective?.Inherits, "CLR page type used first for runtime classification."),
                Row("pageDirective.masterPageFile", artifact.PageDirective?.MasterPageFile, "Master-page declaration preserved as evidence."),
                Row("pageDirective.language", artifact.PageDirective?.Language, "Declared source language."),
                Row("pageDirective.codeBehind", artifact.PageDirective?.CodeBehind, "Code-behind declaration; evidence only and never deployed by this workflow."),
                Row("pageDirective.codeFile", artifact.PageDirective?.CodeFile, "Code-file declaration; evidence only and never deployed by this workflow.")
            });

            writer.Table(
                "Source ASPX Page-directive attributes",
                new[] { "Name", "Value" },
                (artifact.PageDirective?.Attributes ?? Array.Empty<PageDirectiveAttribute>())
                    .Select(attribute => Row(attribute.Name, attribute.Value)));
        }

        private static void AppendIngredientModel(
            MarkdownReportWriter writer,
            PublishingPageCaptureBundle snapshot,
            PublishingPageMigrationPlan plan)
        {
            writer.Table("Ingredient graph projections", new[] { "Property", "Value", "How to read it" }, new[]
            {
                Row("snapshot.ingredientGraph.schemaVersion", snapshot.IngredientGraph.SchemaVersion, "Nested graph contract carried by the immutable source snapshot."),
                Row("snapshot.ingredientGraph.projectionVersion", snapshot.IngredientGraph.ProjectionVersion, "Null means a validated legacy capture-time projection; it does not invalidate or rewrite the source snapshot digest."),
                Row("snapshot.ingredientGraph.nodeCount", snapshot.IngredientGraph.Nodes.Count, "Ingredients understood when the source export was sealed."),
                Row("snapshot.ingredientGraph.edgeCount", snapshot.IngredientGraph.Edges.Count, "Capture-time dependency relationships sealed into the source snapshot."),
                Row("plan.ingredientGraph.schemaVersion", plan.IngredientGraph.SchemaVersion, "Nested graph contract used by this target-specific plan."),
                Row("plan.ingredientGraph.projectionVersion", plan.IngredientGraph.ProjectionVersion, "Current deterministic projection of the unchanged typed source evidence."),
                Row("plan.ingredientGraph.nodeCount", plan.IngredientGraph.Nodes.Count, "Current ingredients for which the plan must choose an action."),
                Row("plan.ingredientGraph.edgeCount", plan.IngredientGraph.Edges.Count, "Current dependency tree evaluated for executability and loss."),
                Row("plan.ingredientGraph.reprojected", !PublishingPageValidationCanonical.Equals(snapshot.IngredientGraph, plan.IngredientGraph), "True means planner capabilities evolved after capture. The source snapshot and its digest remain unchanged; only the plan-time derived graph is newer.")
            });

            writer.Table(
                $"Canonical ingredient nodes — capture-time ({snapshot.IngredientGraph.Nodes.Count})",
                new[] { "ID", "Kind", "Label", "Has content", "Ownership", "Source authority", "Evidence digest", "Runtime requirement", "Evidence references" },
                snapshot.IngredientGraph.Nodes.Select(node => Row(
                    node.Id,
                    node.Kind,
                    node.Label,
                    node.HasContent,
                    node.Ownership,
                    node.SourceAuthority,
                    node.EvidenceDigest,
                    node.RuntimeRequirement,
                    Join(node.EvidenceReferences))));

            writer.Table(
                $"Canonical ingredient edges — capture-time ({snapshot.IngredientGraph.Edges.Count})",
                new[] { "From", "Relationship", "To", "Requirement", "Condition" },
                snapshot.IngredientGraph.Edges.Select(edge => Row(
                    edge.FromIngredientId,
                    edge.Relationship,
                    edge.ToIngredientId,
                    edge.Requirement,
                    edge.Condition)));

            writer.Table(
                $"Plan-time canonical ingredient nodes ({plan.IngredientGraph.Nodes.Count})",
                new[] { "ID", "Kind", "Label", "Has content", "Ownership", "Source authority", "Evidence digest", "Runtime requirement", "Evidence references" },
                plan.IngredientGraph.Nodes.Select(node => Row(
                    node.Id,
                    node.Kind,
                    node.Label,
                    node.HasContent,
                    node.Ownership,
                    node.SourceAuthority,
                    node.EvidenceDigest,
                    node.RuntimeRequirement,
                    Join(node.EvidenceReferences))));

            writer.Table(
                $"Plan-time canonical ingredient edges ({plan.IngredientGraph.Edges.Count})",
                new[] { "From", "Relationship", "To", "Requirement", "Condition" },
                plan.IngredientGraph.Edges.Select(edge => Row(
                    edge.FromIngredientId,
                    edge.Relationship,
                    edge.ToIngredientId,
                    edge.Requirement,
                    edge.Condition)));

            writer.Table(
                $"Ingredient actions ({plan.IngredientActions.Count})",
                new[] { "Action ID", "Ingredient", "Capability", "Disposition", "Realization", "Target identity", "Policy", "Reason", "Released dependencies", "Verification assertions" },
                plan.IngredientActions.Select(action => Row(
                    action.ActionId,
                    action.IngredientId,
                    action.Capability,
                    action.Disposition,
                    action.Realization,
                    action.TargetIdentity,
                    $"{Format(action.PolicyId)}@{Format(action.PolicyVersion)}",
                    action.Reason,
                    Join(action.ReleasedDependencyIngredientIds),
                    Join(action.VerificationAssertions))));

            writer.Table(
                $"Ingredient dependency issues ({plan.IngredientIssues.Count})",
                new[] { "Code", "Severity", "Subject", "Ingredient", "Message", "Source identity", "Target identity" },
                plan.IngredientIssues.Select(issue => Row(
                    issue.Code,
                    issue.Severity,
                    issue.Subject,
                    issue.Ingredient,
                    issue.Message,
                    issue.SourceIdentity,
                    issue.TargetIdentity)));
        }

        private static void AppendFields(MarkdownReportWriter writer, PublishingPageCaptureBundle snapshot, PublishingPageMigrationPlan plan)
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
                    Row("taxonomyBinding", field.TaxonomyBinding == null ? null : $"field={field.TaxonomyBinding.FieldId:D}/{Format(field.TaxonomyBinding.FieldInternalName)}; store={field.TaxonomyBinding.TermStoreId:D}; boundSet={field.TaxonomyBinding.BoundTermSetId:D}; textField={field.TaxonomyBinding.TextFieldId:D}; open={field.TaxonomyBinding.Open}"),
                    Row("taxonomyValueSetSha256", field.TaxonomyValueSetSha256),
                    Row("binaryBase64", Summarize(field.BinaryBase64)),
                    Row("rawType", field.RawType),
                    Row("rawValue", Summarize(field.RawValue)),
                    Row("rawValueJson", Summarize(field.RawValueJson)),
                    Row("captureStatus", field.CaptureStatus),
                    Row("diagnostics", Join(field.Diagnostics))
                });
            }

            var relationshipActions = plan.TaxonomyRelationshipActions.ToDictionary(
                value => TaxonomyRelationshipKey(value.SourceFieldId, value.SourceTermId, value.SourceWssId),
                StringComparer.Ordinal);
            writer.Table("Taxonomy relationship evidence and target actions",
                new[] { "Field / Term", "Source relationship", "Live Term evidence", "Value hidden row", "TaxCatchAll hidden row", "Proof", "Target action", "Target identities", "Verification / reason" },
                snapshot.Fields
                    .Where(field => field.Kind == PageFieldValueKind.Taxonomy || field.Kind == PageFieldValueKind.TaxonomyCollection)
                    .SelectMany(field => field.TaxonomyValues.Select(value =>
                    {
                        Guid termId;
                        Guid.TryParse(value.TermGuid, out termId);
                        relationshipActions.TryGetValue(TaxonomyRelationshipKey(field.Id, termId, value.WssId), out var action);
                        var relationship = value.Relationship;
                        return Row(
                            $"{field.InternalName}; fieldId={field.Id:D}; term={Format(value.TermGuid)}; sourceWssId={value.WssId}; label={Format(value.Label)}",
                            relationship?.State,
                            relationship == null ? null : $"set={Format(relationship.LiveTermSetId)}; setName={Format(relationship.LiveTermSetName)}; label={Format(relationship.LiveTermLabel)}; path={Format(relationship.LiveTermPath)}; taggable={Format(relationship.LiveTermAvailableForTagging)}",
                            FormatHiddenListEntry(relationship?.ValueHiddenListEntry, value.Label),
                            FormatHiddenListEntry(relationship?.TaxCatchAllHiddenListEntry, value.Label),
                            relationship == null ? null : $"fieldValueSetSha256={Format(relationship.SourceFieldValueSetSha256)}; evidenceSha256={Format(relationship.EvidenceSha256)}; capturedAt={relationship.CapturedAtUtc:O}; diagnostics={Join(relationship.Diagnostics)}",
                            action?.Disposition,
                            action == null
                                ? null
                                : action.Disposition == TaxonomyRelationshipDisposition.RetainEvidenceOnly
                                    ? "none; sealed source evidence only"
                                    : $"field={action.TargetFieldId:D}; textField={action.TargetTextFieldId:D}; open={Format(action.TargetFieldOpen)}; store={action.TargetTermStoreId:D}; boundSet={action.TargetBoundTermSetId:D}; liveSet={Format(action.TargetLiveTermSetId)}; valueHiddenSet={Format(action.TargetValueHiddenListTermSetId)}; taxCatchAllSet={Format(action.TargetTaxCatchAllHiddenListTermSetId)}",
                            action == null ? null : $"assertions={Join(action.VerificationAssertions)}; reason={Format(action.Reason)}");
                    })));
        }

        private static void AppendDependencies(MarkdownReportWriter writer, PublishingPageCaptureBundle snapshot, PublishingPageMigrationPlan plan)
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
                snapshot.Security.RoleAssignments.Select(item => Row(item.PrincipalLoginName, item.PrincipalTitle, Join(item.RoleDefinitionNames))));
        }

        private static void AppendPlan(MarkdownReportWriter writer, PublishingPageMigrationPlan plan)
        {
            writer.Table("Plan and target probe", new[] { "Property", "Value", "How to read it" }, new[]
            {
                Row("plan.operation", plan.Operation, "CreatePage is executable now; ApplyDeferredFields reserves future snapshot-based recovery."),
                Row("plan.sourceSnapshotDigest", plan.SourceSnapshotDigest, "Binds this plan to the exact source snapshot."),
                Row("plan.sourceWebUrl", plan.SourceWebUrl, "Source web used for mapping."),
                Row("plan.sourcePageServerRelativeUrl", plan.SourcePageServerRelativeUrl, "Source page represented by the snapshot."),
                Row("plan.originalIdentifier", plan.OriginalIdentifier, "Stable source-qualified Site/Web/file identity written to the target page for ownership verification."),
                Row("plan.targetWebUrl", plan.TargetWebUrl, "Target connection must point here during import."),
                Row("plan.targetWebServerRelativeUrl", plan.TargetWebServerRelativeUrl, "Target web boundary for paths."),
                Row("plan.preferredTargetPageServerRelativeUrl", plan.PreferredTargetPageServerRelativeUrl, "Path produced by exact relative-path mapping before live collision resolution."),
                Row("plan.targetPageServerRelativeUrl", plan.TargetPageServerRelativeUrl, "Final sealed target page path. Only the colliding file leaf may differ from the preferred path."),
                Row("plan.targetPathCollisionResolved", plan.TargetPathCollisionResolved, "True when planning moved only the page filename to avoid a proven foreign collision."),
                Row("plan.targetPathResolutionReason", plan.TargetPathResolutionReason, "Evidence explaining why the final target path differs from the preferred path."),
                Row("plan.pageLayoutName", plan.PageLayoutName, "Publishing layout name passed to page creation."),
                Row("plan.createOnly", plan.CreateOnly, "Import never overwrites an existing object. Planning resolves a foreign collision to another leaf; a post-approval collision invalidates the plan."),
                Row("planningPolicy.targetPageServerRelativeUrl", plan.PlanningPolicy.TargetPageServerRelativeUrl, "Normalized requested target page."),
                Row("planningPolicy.requireInheritedPermissions", plan.PlanningPolicy.RequireInheritedPermissions, "Blocks source pages with unique permissions when true."),
                Row("planningPolicy.blockOnManagedMetadata", plan.PlanningPolicy.BlockOnManagedMetadata, "Legacy compatibility input. It cannot bypass relationship evidence, reviewed mappings, target-state admission, or the no-Term-repair invariant."),
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
                Row("targetProbe.pageContentTypeId", plan.TargetProbe.PageContentTypeId, "Exact Pages-library Content Type ID sealed into the plan and verified after creation."),
                Row("targetProbe.pageLayoutUrl", plan.TargetProbe.PageLayoutUrl, "Approved Publishing Page Layout found or planned in the target site collection."),
                Row("targetProbe.pageLayoutExists", plan.TargetProbe.PageLayoutExists, "False is a blocker."),
                Row("targetProbe.preferredTargetPageServerRelativeUrl", plan.TargetProbe.PreferredTargetPageServerRelativeUrl, "Exact relative-path target inspected before collision allocation."),
                Row("targetProbe.targetPageServerRelativeUrl", plan.TargetProbe.TargetPageServerRelativeUrl, "Final page path sealed by the planning probe."),
                Row("targetProbe.preferredTargetPageExists", plan.TargetProbe.PreferredTargetPageExists, "When true during planning, a foreign preferred-path collision must be resolved without overwriting it."),
                Row("targetProbe.targetPathCollisionResolved", plan.TargetProbe.TargetPathCollisionResolved, "Whether the probe allocated a stable suffix at the page leaf."),
                Row("targetProbe.targetPathResolutionReason", plan.TargetProbe.TargetPathResolutionReason, "Collision evidence retained for review."),
                Row("targetProbe.targetPageExists", plan.TargetProbe.TargetPageExists, "True means the final sealed create-only path was occupied at this probe boundary."),
                Row("targetProbe.existingDependencyPaths", Join(plan.TargetProbe.ExistingDependencyPaths), "Create-only dependency collisions.")
            });

            writer.Table("Approved taxonomy schema mappings",
                new[] { "Source term store", "Source term set", "Target term store", "Target term set", "Interpretation" },
                plan.PlanningPolicy.TaxonomySchemaMappings.Select(item => Row(
                    item.SourceTermStoreId,
                    item.SourceTermSetId,
                    item.TargetTermStoreId,
                    item.TargetTermSetId,
                    "This exact source store/set identity maps to the reviewed target pair for Page Layout schema and page taxonomy-relationship planning.")));

            PublishingPageLayoutPlanReportSection.Append(writer, plan);
        }

        private static string DisplayWorkflow(string value)
        {
            return string.Equals(value, "enterprise-wiki-v1", StringComparison.Ordinal)
                ? "Enterprise Wiki v1"
                : Format(value);
        }

        private static string[] Row(params object[] values) => PublishingPageReportValueFormatter.Row(values);

        private static string Format(object value) => PublishingPageReportValueFormatter.Format(value);

        private static string Join(System.Collections.Generic.IEnumerable<string> values) => PublishingPageReportValueFormatter.Join(values);

        private static string Summarize(string value) => PublishingPageReportValueFormatter.SummarizePayload(value);

        private static string TaxonomyRelationshipKey(Guid fieldId, Guid termId, int sourceWssId)
        {
            return fieldId.ToString("D") + "/" + termId.ToString("D") + "/" + sourceWssId;
        }

        private static string FormatHiddenListEntry(PnP.Framework.Migration.Taxonomy.TaxonomyHiddenListEntrySnapshot entry, string capturedLabel)
        {
            return entry == null
                ? null
                : $"wssId={entry.WssId}; store={entry.TermStoreId:D}; set={entry.TermSetId:D}; term={entry.TermId:D}; title={Format(entry.Title)}; termText={Format(entry.PreferredTerm(capturedLabel))}; path={Format(entry.PreferredPath(capturedLabel))}; catchAllData={Format(entry.CatchAllData)}; catchAllDataLabel={Format(entry.CatchAllDataLabel)}";
        }
    }
}
