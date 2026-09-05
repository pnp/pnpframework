using PnP.Framework.Migration.Evidence;
using PnP.Framework.Migration.Features;
using PnP.Framework.Migration.Diagnostics;
using PnP.Framework.Migration.Lists.Capture;
using PnP.Framework.Migration.Lists.Fields;
using PnP.Framework.Migration.Lists.Items;
using PnP.Framework.Migration.Lists.Planning;
using PnP.Framework.Migration.Lists.Views;
using PnP.Framework.Migration.Pages.Assessment;
using PnP.Framework.Migration.Pages.Ingredients;
using PnP.Framework.Migration.Pages.Publishing.Ingredients;
using PnP.Framework.Migration.Schema.ContentTypes;
using PnP.Framework.Migration.Schema.Fields;
using System;
using System.Collections.Generic;
using System.Linq;

namespace PnP.Framework.Migration.Pages.Publishing.Assessment
{
    internal static class PublishingPageListAssessmentProjector
    {
        private static readonly HashSet<string> ListObjectIssueCodes = new HashSet<string>(
            new[]
            {
                "ListEvidenceUnavailable",
                "UnsupportedListTemplate",
                "ListItemCaptureIncomplete",
                "CalculatedFieldDependencyCycle"
            },
            StringComparer.Ordinal);

        public static void Project(
            PublishingPageAssessmentContext context,
            PublishingPageAssessmentAccumulator assessments)
        {
            AddSharedContentTypeClosures(context, assessments);
            var plans = (context.ListPlan?.Lists ?? Array.Empty<ListMaterializationPlan>())
                .Where(value => value != null)
                .GroupBy(value => value.SourceListId)
                .ToDictionary(group => group.Key, group => group.First());
            foreach (var source in context.Snapshot.ListDependencies.Where(value => value != null))
            {
                plans.TryGetValue(source.SourceListId, out var plan);
                AddList(context, source, plan, assessments);
                AddFields(source, plan, assessments);
                AddContentTypes(source, plan, assessments);
                AddItems(source, plan, assessments);
                AddViewRenderingResources(source, plan, assessments);
                AddViews(source, plan, assessments);
                AddPlatformFeatures(source, plan, assessments);
            }
        }

        private static void AddList(
            PublishingPageAssessmentContext context,
            ListDependencySnapshot source,
            ListMaterializationPlan plan,
            PublishingPageAssessmentAccumulator assessments)
        {
            var unavailable = source.Availability is EvidenceAvailability.Unavailable
                or EvidenceAvailability.Conflict;
            var issues = plan?.Issues ?? Array.Empty<MigrationIssue>();
            var issue = issues.FirstOrDefault(value =>
                (value.Severity is MigrationIssueSeverity.Blocker or MigrationIssueSeverity.Error)
                && ListObjectIssueCodes.Contains(value.Code));
            var blocked = unavailable
                || plan == null
                || issue != null;
            assessments.Add(
                PublishingPageIngredientIds.List(source.SourceWebId, source.SourceListId),
                blocked
                    ? PageIngredientAssessmentState.KnownGap
                    : PageIngredientAssessmentState.TargetInspectionRequired,
                blocked
                    ? unavailable ? IngredientCapability.Missing : IngredientCapability.Incompatible
                    : IngredientCapability.Available,
                blocked ? IngredientDisposition.Defer : IngredientDisposition.Preserve,
                blocked ? "none" : "create-or-reuse-owned-list",
                "policy.list.dependency",
                unavailable
                    ? "Source List evidence is unavailable or conflicting."
                    : plan == null
                        ? context.ListPlanningFailure ?? "No source-authoritative List plan was produced."
                        : issue != null
                            ? issue.Code + ": " + issue.Message
                            : plan.Disposition == ListMaterializationDisposition.Block
                                ? "The List object has a source-authoritative exact-path plan. Child field, Content Type, item, document, View, and feature ingredients retain their own independent pending actions."
                                : "Materialize or reuse the captured List dependency at its exact mapped relative path.",
                blocked ? null : plan.TargetRootFolderServerRelativeUrl,
                blocked
                    ? unavailable ? "ListEvidenceUnavailable" : issue?.Code ?? "ListPlanUnavailable"
                    : null,
                blocked ? null : "Fresh target inspection selects create, owned reuse, recovery, or a suffix only at an observed foreign collision.",
                blocked ? null : "Fresh readback verifies List identity, schema, ownership, and provenance.");
        }

        private static void AddFields(
            ListDependencySnapshot source,
            ListMaterializationPlan plan,
            PublishingPageAssessmentAccumulator assessments)
        {
            var plans = (plan?.Fields ?? Array.Empty<ListFieldMaterializationPlan>())
                .Where(value => value != null)
                .GroupBy(value => value.SourceFieldId)
                .ToDictionary(group => group.Key, group => group.First());
            foreach (var field in source.Fields.Where(value => value != null))
            {
                plans.TryGetValue(field.Id, out var fieldPlan);
                var unavailable = field.Availability is EvidenceAvailability.Unavailable
                    or EvidenceAvailability.Conflict;
                var blocked = unavailable
                    || fieldPlan == null
                    || fieldPlan.Disposition == ListFieldMaterializationDisposition.Block;
                var evidenceOnly = !blocked
                    && fieldPlan.Disposition == ListFieldMaterializationDisposition.EvidenceOnly;
                var mapping = blocked
                    ? (IngredientCapability.Incompatible, IngredientDisposition.Defer, "none")
                    : Map(fieldPlan.Disposition);
                assessments.Add(
                    PublishingPageIngredientIds.ListField(source.SourceWebId, source.SourceListId, field.Id),
                    blocked
                        ? PageIngredientAssessmentState.KnownGap
                        : evidenceOnly
                            ? PageIngredientAssessmentState.Determined
                            : PageIngredientAssessmentState.TargetInspectionRequired,
                    unavailable ? IngredientCapability.Missing : mapping.Item1,
                    mapping.Item2,
                    mapping.Item3,
                    "policy.list-field." + (fieldPlan?.Disposition.ToString().ToLowerInvariant() ?? "missing"),
                    unavailable
                        ? "Field schema evidence is unavailable or conflicting."
                        : fieldPlan?.Reason ?? "No List field materialization decision was produced.",
                    blocked || evidenceOnly || plan == null
                        ? null
                        : plan.TargetRootFolderServerRelativeUrl + "#field:" + field.InternalName,
                    blocked ? FindIssueCode(plan, field.Id, field.InternalName, "ListFieldMaterializationUnavailable") : null,
                    evidenceOnly
                        ? "The omitted schema and captured raw evidence remain in the immutable snapshot."
                        : blocked ? null : $"Fresh target inspection verifies the approved schema policy for field '{field.InternalName}'.");
            }
        }

        private static void AddContentTypes(
            ListDependencySnapshot source,
            ListMaterializationPlan plan,
            PublishingPageAssessmentAccumulator assessments)
        {
            var capturedParents = new HashSet<string>(
                source.SiteContentTypes.Where(value => value != null).Select(value => value.ContentTypeId),
                StringComparer.OrdinalIgnoreCase);
            foreach (var contentType in source.ContentTypes.Where(value => value != null))
            {
                var missingParent = !string.IsNullOrWhiteSpace(contentType.ParentId)
                    && !ContentTypeRuntimeCatalog.IsTargetRuntime(contentType.ParentId)
                    && !capturedParents.Contains(contentType.ParentId);
                assessments.Add(
                    PublishingPageIngredientIds.ListContentType(source.SourceWebId, source.SourceListId, contentType.Id),
                    missingParent
                        ? PageIngredientAssessmentState.KnownGap
                        : PageIngredientAssessmentState.TargetInspectionRequired,
                    missingParent ? IngredientCapability.Missing : IngredientCapability.Available,
                    missingParent ? IngredientDisposition.Defer : IngredientDisposition.Preserve,
                    missingParent ? "none" : "materialize-list-content-type-membership",
                    "policy.list-content-type.membership",
                    missingParent
                        ? "The List-local Content Type references a custom parent whose exact site Content Type closure is absent."
                        : "Create or reuse the captured List Content Type membership, field links, and ordering after its parent closure is admitted.",
                    missingParent || plan == null
                        ? null
                        : plan.TargetRootFolderServerRelativeUrl + "#content-type:" + contentType.Id,
                    missingParent ? "CustomListContentTypeClosureUnavailable" : null,
                    missingParent ? null : $"The List receipt maps source Content Type '{contentType.Id}' to a verified target Content Type ID.");
            }
        }

        private static void AddItems(
            ListDependencySnapshot source,
            ListMaterializationPlan plan,
            PublishingPageAssessmentAccumulator assessments)
        {
            var fields = (plan?.Fields ?? Array.Empty<ListFieldMaterializationPlan>())
                .Where(value => value != null)
                .GroupBy(value => value.InternalName, StringComparer.OrdinalIgnoreCase)
                .ToDictionary(group => group.Key, group => group.First(), StringComparer.OrdinalIgnoreCase);
            foreach (var item in source.Items.Where(value => value != null))
            {
                var unavailable = item.Availability is EvidenceAvailability.Unavailable
                    or EvidenceAvailability.Conflict;
                var snapshotOnlyValues = new List<string>();
                var transformed = false;
                foreach (var value in item.Values.Where(value => value != null && value.Kind != ListItemValueKind.Null))
                {
                    if (string.Equals(value.InternalName, "ContentTypeId", StringComparison.OrdinalIgnoreCase))
                    {
                        transformed = true;
                        continue;
                    }
                    if (!fields.TryGetValue(value.InternalName, out var fieldPlan)
                        || fieldPlan.Disposition is ListFieldMaterializationDisposition.Block
                            or ListFieldMaterializationDisposition.EvidenceOnly)
                    {
                        snapshotOnlyValues.Add(value.InternalName);
                        transformed = true;
                        continue;
                    }
                    if (fieldPlan.Disposition is ListFieldMaterializationDisposition.MapLookup
                        or ListFieldMaterializationDisposition.MapTaxonomy
                        or ListFieldMaterializationDisposition.RequireTargetRuntime
                        or ListFieldMaterializationDisposition.CreateOrReuseOwnedCalculated
                        or ListFieldMaterializationDisposition.CreateOrReuseOwnedSchemaOnly)
                    {
                        transformed = true;
                    }
                }

                var blocked = unavailable || plan == null;
                assessments.Add(
                    PublishingPageIngredientIds.ListItem(source.SourceWebId, source.SourceListId, item.SourceItemId),
                    blocked
                        ? PageIngredientAssessmentState.KnownGap
                        : PageIngredientAssessmentState.TargetInspectionRequired,
                    blocked
                        ? unavailable || plan == null ? IngredientCapability.Missing : IngredientCapability.Incompatible
                        : IngredientCapability.Available,
                    blocked
                        ? IngredientDisposition.Defer
                        : transformed ? IngredientDisposition.Transform : IngredientDisposition.Preserve,
                    blocked
                        ? "none"
                        : snapshotOnlyValues.Count > 0
                            ? "replay-recognized-values-retain-others-snapshot-only"
                            : transformed
                                ? "replay-approved-values-and-substitute-runtime-values"
                                : "replay-approved-current-values",
                    "policy.list-item.current-state",
                    unavailable
                        ? "The captured List item evidence is unavailable or conflicting."
                        : plan == null
                            ? "The owning List has no source-authoritative target path plan."
                            : snapshotOnlyValues.Count > 0
                                ? "Replay recognized values and retain unrecognized or intentionally omitted fields only in the immutable snapshot: " + string.Join(", ", snapshotOnlyValues.Distinct(StringComparer.OrdinalIgnoreCase)) + "."
                                : transformed
                                    ? "Replay recognized values while remapping identity-bound values and regenerating reviewed target-runtime values."
                                    : "Replay every nonempty captured value through an approved lossless field action.",
                    blocked ? null : plan.TargetRootFolderServerRelativeUrl + "#source-item:" + item.SourceItemId,
                    blocked
                        ? unavailable ? "ListItemEvidenceUnavailable" : "ListPlanUnavailable"
                        : null,
                    blocked ? null : $"The List receipt contains a source-to-target item ID mapping for source item '{item.SourceItemId}'.",
                    blocked ? null : snapshotOnlyValues.Count > 0
                        ? "Fresh readback verifies every approved value and the item provenance digest; snapshot-only fields are not fabricated on the target."
                            : "Fresh readback verifies every approved value and the item provenance digest.");

                if (item.Document != null)
                {
                    AddDocument(source, plan, item, assessments);
                    if (item.Document.InformationProtection != null)
                    {
                        AddInformationProtection(source, plan, item, assessments);
                    }
                }
                foreach (var attachment in item.Attachments.Where(value => value != null))
                {
                    AddAttachment(source, plan, item, attachment, assessments);
                }
            }
        }

        private static void AddDocument(
            ListDependencySnapshot source,
            ListMaterializationPlan plan,
            ListItemSnapshot item,
            PublishingPageAssessmentAccumulator assessments)
        {
            var document = item.Document;
            var bytesMissing = document.Kind == ListDocumentObjectKind.File
                && !ListMigrationPlanFactory.HasReplayableBinary(document.Content);
            var archived = bytesMissing && ListMigrationPlanFactory.IsArchivedContent(document.Content);
            var rightsManaged = document.Kind == ListDocumentObjectKind.File
                && ListMigrationPlanFactory.IsRightsManagedEnvelope(document.Content);
            var informationProtected = document.InformationProtection != null;
            var unclassified = document.Kind == ListDocumentObjectKind.File
                && ListMigrationPlanFactory.IsUnclassifiedBinary(document.Content);
            var deferred = bytesMissing || rightsManaged || unclassified || plan == null;
            var targetPath = plan == null
                ? null
                : MapListOwnedPath(source, plan, document.ServerRelativeUrl);
            assessments.Add(
                PublishingPageIngredientIds.ListDocument(source.SourceWebId, source.SourceListId, item.SourceItemId),
                deferred
                    ? PageIngredientAssessmentState.KnownGap
                    : PageIngredientAssessmentState.TargetInspectionRequired,
                deferred
                    ? rightsManaged || unclassified ? IngredientCapability.Unknown : IngredientCapability.Missing
                    : IngredientCapability.Available,
                deferred ? IngredientDisposition.Defer : IngredientDisposition.Preserve,
                deferred
                    ? archived
                        ? "reactivate-source-and-recapture"
                        : rightsManaged
                            ? "retain-envelope-and-logical-identity-pending-replay-evidence"
                            : unclassified
                                ? "fresh-capture-to-classify-binary-representation"
                            : "none"
                    : document.Kind == ListDocumentObjectKind.Folder
                        ? "create-or-reuse-folder"
                        : "copy-exact-bytes-create-only",
                "policy.list-document.current-state",
                archived
                    ? "The source document is stored in Microsoft 365 Archive. Its exact bytes remain unavailable until source reactivation followed by a fresh capture."
                    : rightsManaged
                    ? informationProtected
                        ? "SharePoint returned a protected envelope for document-level Information Protection label '"
                            + document.InformationProtection.LabelId
                            + "'. Repeated reads keep source UniqueId, version, logical length, envelope length, cTag, and QuickXorHash stable while changing the envelope SHA. Source library IRM is "
                            + (source.InformationRightsManagement == null
                                ? "not captured"
                                : source.InformationRightsManagement.IrmEnabled ? "enabled" : "disabled")
                            + "; cross-site protected-payload replay and semantic verification have not been proven."
                        : "SharePoint returned an Information Rights Management envelope whose payload SHA changes across repeated CSOM and REST reads while source UniqueId, version, length, cTag, and QuickXorHash remain stable. The exact returned envelope is retained, but cross-site replay and semantic verification have not been proven."
                    : unclassified
                    ? "This immutable legacy snapshot predates binary representation classification. Exact returned bytes remain sealed, but ordinary-file versus rights-managed-envelope replay semantics are not yet proven."
                    : bytesMissing
                    ? "Exact current document bytes are absent from the sealed snapshot."
                    : plan == null
                        ? "The owning List has no source-authoritative target path plan."
                        : "Materialize the captured current document or folder under the exact mapped List path.",
                targetPath,
                deferred
                    ? bytesMissing
                        ? archived ? "ListBinaryContentArchived" : "ListBinaryEvidenceUnavailable"
                        : rightsManaged
                            ? informationProtected
                                ? "ListInformationProtectedBinaryReplayUnverified"
                                : "ListRightsManagedBinaryReplayUnverified"
                            : unclassified
                                ? "ListBinaryRepresentationUnclassified"
                                : "ListPlanUnavailable"
                    : null,
                archived
                    ? "Before target execution, source reactivation plus fresh capture must retain the same source item and file identity and supply a replayable current payload."
                    : rightsManaged
                        ? "An explicitly approved canary uploads the retained returned envelope as an opaque payload; it does not decrypt, relabel, or repair the document."
                        : deferred
                            ? null
                    : document.Kind == ListDocumentObjectKind.Folder
                        ? "Fresh readback verifies the target folder path and provenance."
                        : $"Fresh readback verifies target file bytes with SHA-256 '{document.Content?.Artifact?.Sha256}'.",
                rightsManaged
                    ? "Fresh target readback verifies the exact planned path, migrated item provenance, protected-envelope representation, and logical file length '"
                        + document.Length + "'."
                    : null,
                rightsManaged
                    ? "Returned envelope SHA-256 and publish-license SHA-256 are diagnostic rather than equality assertions because repeated source reads proved both values are nondeterministic."
                    : null,
                rightsManaged && informationProtected
                    ? "Fresh readback records the observed Information Protection label relationship and access/decryption behavior without substituting a target-tenant label."
                    : null);
        }

        private static void AddInformationProtection(
            ListDependencySnapshot source,
            ListMaterializationPlan plan,
            ListItemSnapshot item,
            PublishingPageAssessmentAccumulator assessments)
        {
            var informationProtection = item.Document.InformationProtection;
            var targetPath = plan == null
                ? null
                : MapListOwnedPath(source, plan, item.Document.ServerRelativeUrl);
            assessments.Add(
                PublishingPageIngredientIds.ListDocumentInformationProtection(
                    source.SourceWebId,
                    source.SourceListId,
                    item.SourceItemId),
                PageIngredientAssessmentState.KnownGap,
                IngredientCapability.Unknown,
                IngredientDisposition.Defer,
                "retain-source-label-evidence-and-run-approved-opaque-envelope-canary",
                "policy.list-document-information-protection.current-state",
                "The source document is assigned Information Protection label '"
                    + informationProtection.LabelId
                    + "' with assignment method '" + (informationProtection.AssignmentMethod ?? string.Empty)
                    + "'. The exact relationship is retained, but the target tenant label identity, cryptographic protection boundary, and user access semantics have not been verified.",
                targetPath,
                "ListDocumentInformationProtectionReplayUnverified",
                "Fresh target readback records `_IpLabelId`, assignment method, label hash, promotion cTag, and decrypt-skip evidence when exposed by SharePoint.",
                "The observed label relationship is compared with source label '"
                    + informationProtection.LabelId
                    + "'; any mismatch remains explicit and is not repaired or mapped to a similar target label.",
                "Target access and decryption behavior are reported independently from byte transport success.");
        }

        private static void AddAttachment(
            ListDependencySnapshot source,
            ListMaterializationPlan plan,
            ListItemSnapshot item,
            ListAttachmentSnapshot attachment,
            PublishingPageAssessmentAccumulator assessments)
        {
            var bytesMissing = !ListMigrationPlanFactory.HasReplayableBinary(attachment.Content);
            var archived = bytesMissing && ListMigrationPlanFactory.IsArchivedContent(attachment.Content);
            var unclassified = ListMigrationPlanFactory.IsUnclassifiedBinary(attachment.Content);
            var blocked = bytesMissing || unclassified || plan == null;
            assessments.Add(
                PublishingPageIngredientIds.ListAttachment(source.SourceWebId, source.SourceListId, item.SourceItemId, attachment.FileName),
                blocked
                    ? PageIngredientAssessmentState.KnownGap
                    : PageIngredientAssessmentState.TargetInspectionRequired,
                blocked
                    ? unclassified ? IngredientCapability.Unknown : IngredientCapability.Missing
                    : IngredientCapability.Available,
                blocked ? IngredientDisposition.Defer : IngredientDisposition.Preserve,
                blocked
                    ? archived
                        ? "reactivate-source-and-recapture"
                        : unclassified
                            ? "fresh-capture-to-classify-binary-representation"
                            : "none"
                    : "copy-exact-bytes-create-only",
                "policy.list-attachment.current-state",
                archived
                    ? "The source attachment is stored in Microsoft 365 Archive. Its exact bytes remain unavailable until source reactivation followed by a fresh capture."
                    : unclassified
                    ? "This immutable legacy snapshot predates binary representation classification. Exact returned bytes remain sealed, but ordinary-file versus rights-managed-envelope replay semantics are not yet proven."
                    : bytesMissing
                    ? "Exact attachment bytes are absent from the sealed snapshot."
                    : plan == null
                        ? "The owning List has no source-authoritative target path plan."
                        : "Copy the exact captured attachment bytes to the materialized target item.",
                blocked ? null : plan.TargetRootFolderServerRelativeUrl + "#source-item:" + item.SourceItemId + "/attachment:" + attachment.FileName,
                blocked
                    ? bytesMissing
                        ? archived ? "ListBinaryContentArchived" : "ListBinaryEvidenceUnavailable"
                        : unclassified ? "ListBinaryRepresentationUnclassified" : "ListPlanUnavailable"
                    : null,
                blocked ? null : $"Fresh readback verifies attachment bytes with SHA-256 '{attachment.Content?.Artifact?.Sha256}'.");
        }

        private static void AddViews(
            ListDependencySnapshot source,
            ListMaterializationPlan plan,
            PublishingPageAssessmentAccumulator assessments)
        {
            var plans = (plan?.Views ?? Array.Empty<ListViewMaterializationPlan>())
                .Where(value => value != null)
                .GroupBy(value => value.SourceViewId)
                .ToDictionary(group => group.Key, group => group.First());
            foreach (var view in source.Views.Where(value => value != null))
            {
                plans.TryGetValue(view.Id, out var viewPlan);
                var unavailable = view.Availability is EvidenceAvailability.Unavailable
                    or EvidenceAvailability.Conflict;
                var customRendering = IsCustomRenderingReference(view.JsLink)
                    || IsCustomRenderingReference(view.XslLink);
                var blocked = unavailable
                    || viewPlan == null
                    || viewPlan.Disposition == ListViewMaterializationDisposition.Block;
                var personal = !blocked && viewPlan.Disposition == ListViewMaterializationDisposition.SkipPersonal;
                assessments.Add(
                    PublishingPageIngredientIds.View(source.SourceWebId, source.SourceListId, view.Id),
                    blocked
                        ? PageIngredientAssessmentState.KnownGap
                        : personal
                            ? PageIngredientAssessmentState.Determined
                            : PageIngredientAssessmentState.TargetInspectionRequired,
                    blocked
                        ? unavailable || viewPlan == null ? IngredientCapability.Missing : IngredientCapability.Incompatible
                        : IngredientCapability.Available,
                    blocked
                        ? IngredientDisposition.Defer
                        : personal ? IngredientDisposition.Drop : IngredientDisposition.Preserve,
                    blocked
                        ? "none"
                        : personal ? "omit-personal-view" : "create-or-reuse-view",
                    "policy.list-view.dependency",
                    unavailable
                        ? "View schema evidence is unavailable or conflicting."
                        : viewPlan?.Reason ?? "No View materialization decision was produced.",
                    blocked || personal || plan == null
                        ? null
                        : plan.TargetRootFolderServerRelativeUrl + "#view:" + view.Id.ToString("D"),
                    blocked
                        ? customRendering ? "ViewRenderingResourceUnavailable" : "ViewEvidenceUnavailable"
                        : null,
                    personal
                        ? "The omitted personal View remains fully represented in the source snapshot and reviewed action."
                        : blocked ? null : "Fresh target inspection and readback verify View identity, schema, and Web Part binding suitability.");
            }
        }

        private static void AddViewRenderingResources(
            ListDependencySnapshot source,
            ListMaterializationPlan plan,
            PublishingPageAssessmentAccumulator assessments)
        {
            var plans = (plan?.ViewRenderingResources ?? Array.Empty<ListViewRenderingResourceMaterializationPlan>())
                .Where(value => value != null && !string.IsNullOrWhiteSpace(value.SourceResourceId))
                .GroupBy(value => value.SourceResourceId, StringComparer.Ordinal)
                .ToDictionary(group => group.Key, group => group.First(), StringComparer.Ordinal);
            foreach (var resource in source.ViewRenderingResources.Where(value => value != null))
            {
                plans.TryGetValue(resource.Id ?? string.Empty, out var resourcePlan);
                var unavailable = resource.Availability is EvidenceAvailability.Unavailable
                    or EvidenceAvailability.Conflict
                    || resource.Artifact == null;
                var blocked = resourcePlan == null
                    || resourcePlan.Disposition == ListViewRenderingResourceMaterializationDisposition.Block;
                var referenceOnly = !blocked
                    && resourcePlan.Disposition == ListViewRenderingResourceMaterializationDisposition.PreserveReferenceOnly;
                assessments.Add(
                    PublishingPageIngredientIds.ViewRenderingResource(source.SourceSiteId, resource.Id),
                    blocked
                        ? PageIngredientAssessmentState.KnownGap
                        : PageIngredientAssessmentState.TargetInspectionRequired,
                    blocked
                        ? unavailable ? IngredientCapability.Missing : IngredientCapability.Incompatible
                        : IngredientCapability.Available,
                    blocked
                        ? IngredientDisposition.Defer
                        : referenceOnly ? IngredientDisposition.Substitute : IngredientDisposition.Preserve,
                    blocked
                        ? "none"
                        : referenceOnly ? "preserve-mapped-reference-without-resource-bytes" : "copy-exact-bytes-create-only",
                    "policy.list-view.rendering-resource",
                    resourcePlan?.Reason
                        ?? (unavailable
                            ? "Exact View rendering-resource bytes are absent from the sealed snapshot."
                            : "No View rendering-resource materialization decision was produced."),
                    blocked ? null : resourcePlan.TargetServerRelativeUrl,
                    blocked ? "ViewRenderingResourceUnavailable" : null,
                    blocked
                        ? null
                        : referenceOnly
                            ? "Fresh target inspection proves the mapped path does not contain a foreign resource that would change the preserved unresolved relationship."
                            : "Fresh target inspection requires an absent path or an exact-byte match; foreign content is a collision.",
                    blocked
                        ? null
                        : referenceOnly
                            ? "Execution creates no resource bytes; fresh View readback verifies the captured reference is retained."
                            : $"Fresh readback verifies rendering-resource bytes with SHA-256 '{resource.Artifact.Sha256}'.");
            }
        }

        private static void AddPlatformFeatures(
            ListDependencySnapshot source,
            ListMaterializationPlan plan,
            PublishingPageAssessmentAccumulator assessments)
        {
            foreach (var feature in plan?.RequiredFeatures ?? Array.Empty<PlatformFeatureMaterializationPlan>())
            {
                var blocked = feature.Disposition == PlatformFeatureMaterializationDisposition.Block;
                assessments.Add(
                    PublishingPageIngredientIds.PlatformFeature(source.SourceSiteId, feature.FeatureId),
                    blocked
                        ? PageIngredientAssessmentState.KnownGap
                        : PageIngredientAssessmentState.TargetInspectionRequired,
                    blocked ? IngredientCapability.Incompatible : IngredientCapability.Available,
                    blocked ? IngredientDisposition.Defer : IngredientDisposition.Substitute,
                    blocked ? "none" : "ensure-target-runtime-feature-active",
                    "policy.platform-feature.required-runtime",
                    feature.Reason,
                    blocked ? null : feature.TargetWebUrl + "#site-feature:" + feature.FeatureId.ToString("D"),
                    blocked ? "PlatformFeatureMaterializationUnavailable" : null,
                    blocked ? null : "Fresh target inspection proves the feature is active or can be activated after topology materialization.",
                    blocked ? null : "Fresh readback verifies the feature and promised runtime Content Types.");
            }
        }

        private static void AddSharedContentTypeClosures(
            PublishingPageAssessmentContext context,
            PublishingPageAssessmentAccumulator assessments)
        {
            var sourceSchemas = context.Snapshot.ListDependencies
                .Where(value => value != null)
                .SelectMany(value => value.SiteContentTypes ?? Array.Empty<ContentTypeSchemaSnapshot>())
                .Where(value => value != null)
                .GroupBy(value => PublishingPageIngredientIds.SiteContentType(SchemaScope(value), value.ContentTypeId), StringComparer.Ordinal)
                .Select(group => group.First())
                .ToArray();
            var planned = (context.ListPlan?.Lists ?? Array.Empty<ListMaterializationPlan>())
                .SelectMany(value => value.SiteContentTypes ?? Array.Empty<ContentTypeClosureNodePlan>())
                .Where(value => value?.Schema != null)
                .ToArray();
            var sharedFields = new List<SharedSiteFieldAssessmentInput>();
            foreach (var source in sourceSchemas)
            {
                var plan = FindSchemaPlan(source, planned);
                var blocked = IsContentTypeObjectUnavailable(plan?.Schema);
                var childFieldPending = !blocked
                    && plan.Schema.Disposition == ContentTypeMaterializationDisposition.Block;
                var scope = SchemaScope(source);
                assessments.Add(
                    PublishingPageIngredientIds.SiteContentType(scope, source.ContentTypeId),
                    blocked
                        ? PageIngredientAssessmentState.KnownGap
                        : PageIngredientAssessmentState.TargetInspectionRequired,
                    blocked ? IngredientCapability.Incompatible : IngredientCapability.Available,
                    blocked ? IngredientDisposition.Defer : IngredientDisposition.Preserve,
                    blocked
                        ? "none"
                        : plan.Schema.Disposition == ContentTypeMaterializationDisposition.ReuseOwned
                            ? "reuse-owned-site-content-type"
                            : "create-or-reuse-owned-site-content-type",
                    "policy.site-content-type.closure",
                    plan == null
                        ? "No site Content Type closure materialization decision was produced."
                        : childFieldPending
                            ? "The Content Type identity and metadata have a source-authoritative action; blocked child field ingredients remain independently pending."
                            : plan.Schema.Reason,
                    blocked ? null : plan.TargetOwnerWebUrl + "#content-type:" + source.ContentTypeId,
                    blocked ? "SiteContentTypeMaterializationUnavailable" : null,
                    blocked ? null : "Fresh target inspection verifies Content Type identity, parent lineage, metadata, field links, and ownership.");

                var fieldPlans = (plan?.Schema?.Fields ?? Array.Empty<FieldSchemaMaterializationPlan>())
                    .Where(value => value != null)
                    .GroupBy(value => value.FieldId)
                    .ToDictionary(group => group.Key, group => group.First());
                foreach (var field in source.RequiredFieldClosure.Where(value => value != null))
                {
                    fieldPlans.TryGetValue(field.Id, out var fieldPlan);
                    sharedFields.Add(new SharedSiteFieldAssessmentInput
                    {
                        IngredientId = PublishingPageIngredientIds.SiteField(scope, field.Id),
                        Field = field,
                        Plan = fieldPlan,
                        TargetOwnerWebUrl = plan?.TargetOwnerWebUrl
                    });
                }
            }

            foreach (var group in sharedFields
                         .GroupBy(value => value.IngredientId, StringComparer.Ordinal)
                         .OrderBy(value => value.Key, StringComparer.Ordinal))
            {
                AddSharedSiteField(group.Key, group, assessments);
            }
        }

        private static void AddSharedSiteField(
            string ingredientId,
            IEnumerable<SharedSiteFieldAssessmentInput> candidates,
            PublishingPageAssessmentAccumulator assessments)
        {
            var inputs = candidates.ToArray();
            var field = inputs
                .OrderByDescending(value => value.Plan?.Disposition == FieldSchemaMaterializationDisposition.CreateOrReuseOwned)
                .ThenBy(value => value.Field.InternalName, StringComparer.Ordinal)
                .Select(value => value.Field)
                .First();
            var missingPlan = inputs.Any(value => value.Plan == null);
            var blockedPlans = inputs
                .Where(value => value.Plan?.Disposition == FieldSchemaMaterializationDisposition.Block)
                .Select(value => value.Plan)
                .ToArray();
            if (missingPlan || blockedPlans.Length > 0)
            {
                var reasons = blockedPlans.Select(value => value.Reason)
                    .Where(value => !string.IsNullOrWhiteSpace(value))
                    .Distinct(StringComparer.Ordinal)
                    .ToList();
                if (missingPlan)
                {
                    reasons.Add("At least one required Content Type consumer has no site field-schema materialization decision.");
                }
                assessments.Add(
                    ingredientId,
                    PageIngredientAssessmentState.KnownGap,
                    missingPlan ? IngredientCapability.Missing : IngredientCapability.Incompatible,
                    IngredientDisposition.Defer,
                    "none",
                    "policy.site-field",
                    string.Join(" ", reasons),
                    mitigationCode: "SiteFieldMaterializationUnavailable");
                return;
            }

            var createPlans = inputs
                .Select(value => value.Plan)
                .Where(value => value.Disposition == FieldSchemaMaterializationDisposition.CreateOrReuseOwned)
                .ToArray();
            if (createPlans.Length > 0)
            {
                var schemaIdentities = createPlans
                    .Select(value => new
                    {
                        value.InternalName,
                        value.TypeAsString,
                        Digest = value.TargetPortableSchemaSha256 ?? value.SourcePortableSchemaSha256
                    })
                    .Distinct()
                    .ToArray();
                var targetOwners = inputs.Select(value => value.TargetOwnerWebUrl)
                    .Where(value => !string.IsNullOrWhiteSpace(value))
                    .Distinct(StringComparer.OrdinalIgnoreCase)
                    .ToArray();
                if (schemaIdentities.Length != 1 || targetOwners.Length != 1)
                {
                    assessments.Add(
                        ingredientId,
                        PageIngredientAssessmentState.KnownGap,
                        IngredientCapability.Incompatible,
                        IngredientDisposition.Defer,
                        "none",
                        "policy.site-field.shared-object",
                        schemaIdentities.Length != 1
                            ? "Content Type consumers disagree on the portable schema of this shared site field."
                            : "Content Type consumers disagree on the target owner Web of this shared site field.",
                        mitigationCode: schemaIdentities.Length != 1
                            ? "SiteFieldSchemaConflict"
                            : "SiteFieldTargetOwnerConflict");
                    return;
                }

                var selected = createPlans[0];
                assessments.Add(
                    ingredientId,
                    PageIngredientAssessmentState.TargetInspectionRequired,
                    IngredientCapability.Available,
                    IngredientDisposition.Preserve,
                    "create-or-reuse-owned-schema",
                    "policy.site-field.shared-object",
                    "At least one Content Type directly defines this shared site field, so its object-level action is create-or-reuse; inherited consumers bind to the same field after materialization. " + selected.Reason,
                    targetOwners[0] + "#field:" + field.Id.ToString("D"),
                    verificationAssertions: $"Fresh target inspection verifies one owned portable schema for shared field '{field.InternalName}' and every dependent Content Type binds to that field GUID.");
                return;
            }

            var runtimeOwners = inputs.Select(value => value.TargetOwnerWebUrl)
                .Where(value => !string.IsNullOrWhiteSpace(value))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToArray();
            if (runtimeOwners.Length != 1)
            {
                assessments.Add(
                    ingredientId,
                    PageIngredientAssessmentState.KnownGap,
                    IngredientCapability.Incompatible,
                    IngredientDisposition.Defer,
                    "none",
                    "policy.site-field.shared-object",
                    "Inherited Content Type consumers disagree on the target owner Web of this runtime-supplied site field.",
                    mitigationCode: "SiteFieldTargetOwnerConflict");
                return;
            }

            assessments.Add(
                ingredientId,
                PageIngredientAssessmentState.TargetInspectionRequired,
                IngredientCapability.Available,
                IngredientDisposition.Substitute,
                "reuse-target-runtime-schema",
                "policy.site-field.shared-object",
                "Every captured Content Type consumer inherits this field from the target runtime or parent Content Type; do not create or repair it.",
                runtimeOwners[0] + "#field:" + field.Id.ToString("D"),
                verificationAssertions: $"Fresh target inspection verifies exact ID, internal name, and type for runtime field '{field.InternalName}'.");
        }

        private static bool IsContentTypeObjectUnavailable(ContentTypeMaterializationPlan plan)
        {
            if (plan == null)
            {
                return true;
            }

            return plan.Disposition == ContentTypeMaterializationDisposition.Block
                && !(plan.Fields?.Any(value => value?.Disposition == FieldSchemaMaterializationDisposition.Block) ?? false);
        }

        private sealed class SharedSiteFieldAssessmentInput
        {
            public string IngredientId { get; set; }

            public FieldSchemaSnapshot Field { get; set; }

            public FieldSchemaMaterializationPlan Plan { get; set; }

            public string TargetOwnerWebUrl { get; set; }
        }

        private static (IngredientCapability, IngredientDisposition, string) Map(
            ListFieldMaterializationDisposition disposition)
        {
            switch (disposition)
            {
                case ListFieldMaterializationDisposition.RequireTargetRuntime:
                case ListFieldMaterializationDisposition.RequireTargetRuntimeAndCopyValue:
                    return (IngredientCapability.Available, IngredientDisposition.Substitute, "reuse-target-runtime-schema");
                case ListFieldMaterializationDisposition.CreateOrReuseOwnedAndCopyValue:
                    return (IngredientCapability.Available, IngredientDisposition.Preserve, "create-or-reuse-owned-and-copy-values");
                case ListFieldMaterializationDisposition.CreateOrReuseOwnedCalculated:
                    return (IngredientCapability.Available, IngredientDisposition.Preserve, "create-or-reuse-owned-calculated-schema");
                case ListFieldMaterializationDisposition.MapLookup:
                    return (IngredientCapability.Available, IngredientDisposition.Transform, "map-lookup-list-and-item-identities");
                case ListFieldMaterializationDisposition.MapTaxonomy:
                    return (IngredientCapability.Available, IngredientDisposition.Transform, "map-taxonomy-store-and-set");
                case ListFieldMaterializationDisposition.CreateOrReuseOwnedSchemaOnly:
                    return (IngredientCapability.Available, IngredientDisposition.Preserve, "create-or-reuse-owned-schema-only");
                case ListFieldMaterializationDisposition.EvidenceOnly:
                    return (IngredientCapability.Unknown, IngredientDisposition.Drop, "retain-snapshot-only");
                default:
                    return (IngredientCapability.Incompatible, IngredientDisposition.Defer, "none");
            }
        }

        private static string FindIssueCode(
            ListMaterializationPlan plan,
            Guid fieldId,
            string internalName,
            string fallback)
        {
            return plan?.Issues?.FirstOrDefault(value =>
                       !string.IsNullOrWhiteSpace(value.Subject)
                       && (value.Subject.IndexOf(fieldId.ToString("D"), StringComparison.OrdinalIgnoreCase) >= 0
                           || value.Subject.IndexOf("field:" + internalName + ":", StringComparison.OrdinalIgnoreCase) >= 0))?.Code
                ?? fallback;
        }

        private static ContentTypeClosureNodePlan FindSchemaPlan(
            ContentTypeSchemaSnapshot source,
            IEnumerable<ContentTypeClosureNodePlan> plans)
        {
            var candidates = plans.Where(value => string.Equals(
                    value.Schema.ContentTypeId,
                    source.ContentTypeId,
                    StringComparison.OrdinalIgnoreCase))
                .ToArray();
            var sourceScope = SchemaScope(source);
            var exact = candidates.FirstOrDefault(value => string.Equals(
                UrlScope(value.SourceOwnerWebUrl),
                sourceScope,
                StringComparison.OrdinalIgnoreCase));
            return exact ?? (candidates.Length == 1 ? candidates[0] : null);
        }

        private static string SchemaScope(ContentTypeSchemaSnapshot schema)
        {
            if (!string.IsNullOrWhiteSpace(schema?.SourceScope))
            {
                return Uri.TryCreate(schema.SourceScope, UriKind.Absolute, out var absolute)
                    ? NormalizeScope(absolute.AbsolutePath)
                    : NormalizeScope(schema.SourceScope);
            }
            return UrlScope(schema?.SourceWebUrl);
        }

        private static string UrlScope(string value)
        {
            return Uri.TryCreate(value, UriKind.Absolute, out var absolute)
                ? NormalizeScope(absolute.AbsolutePath)
                : NormalizeScope(value);
        }

        private static string NormalizeScope(string value)
        {
            var normalized = Uri.UnescapeDataString(value ?? string.Empty).Replace('\\', '/').TrimEnd('/');
            return normalized.Length == 0 ? "/" : normalized;
        }

        private static string MapListOwnedPath(
            ListDependencySnapshot source,
            ListMaterializationPlan plan,
            string sourcePath)
        {
            var root = source.RootFolderServerRelativeUrl?.TrimEnd('/');
            if (!string.IsNullOrWhiteSpace(sourcePath)
                && !string.IsNullOrWhiteSpace(root)
                && sourcePath.StartsWith(root + "/", StringComparison.OrdinalIgnoreCase))
            {
                return plan.TargetRootFolderServerRelativeUrl.TrimEnd('/') + sourcePath.Substring(root.Length);
            }
            return plan.TargetRootFolderServerRelativeUrl;
        }

        private static bool IsCustomRenderingReference(string value)
        {
            return !string.IsNullOrWhiteSpace(value)
                && (value.IndexOf('/') >= 0 || value.IndexOf('\\') >= 0 || value.StartsWith("~", StringComparison.Ordinal));
        }
    }
}
