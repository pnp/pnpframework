using PnP.Framework.Migration.Lists.Planning;
using PnP.Framework.Migration.Diagnostics;
using PnP.Framework.Migration.Evidence;
using PnP.Framework.Migration.Pages.Ingredients;
using PnP.Framework.Migration.Pages.Publishing.Capture;
using PnP.Framework.Migration.Pages.Publishing.Planning;
using System;
using System.Collections.Generic;
using System.Linq;

namespace PnP.Framework.Migration.Pages.Publishing.Ingredients
{
    internal static class PublishingPageListIngredientActionProjector
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
            PublishingPageCaptureBundle snapshot,
            PublishingPageMigrationPlan plan,
            IDictionary<string, PageIngredientAction> actions)
        {
            Project(snapshot, plan, actions, true);
        }

        public static void Project(
            PublishingPageCaptureBundle snapshot,
            PublishingPageMigrationPlan plan,
            IDictionary<string, PageIngredientAction> actions,
            bool transactionDependencyProjection)
        {
            PublishingPageListSchemaIngredientActionProjector.ProjectSharedClosures(snapshot, plan, actions);
            if (plan.ListMigration == null)
            {
                return;
            }

            var sourceByList = snapshot.ListDependencies.ToDictionary(value => value.SourceListId);
            foreach (var listPlan in plan.ListMigration.Lists)
            {
                if (!sourceByList.TryGetValue(listPlan.SourceListId, out var source))
                {
                    continue;
                }

                var blocked = transactionDependencyProjection
                    ? IsListObjectBlocked(source, listPlan)
                    : !listPlan.IsExecutable;
                PublishingPageIngredientActionFactory.Add(actions, PublishingPageIngredientActionFactory.Create(
                    PublishingPageIngredientIds.List(source.SourceWebId, source.SourceListId),
                    blocked ? IngredientCapability.Incompatible : IngredientCapability.Available,
                    blocked ? IngredientDisposition.Block : IngredientDisposition.Preserve,
                    listPlan.Disposition == ListMaterializationDisposition.ReuseOwned
                        ? "reuse-owned"
                        : blocked ? "none" : "create-owned",
                    "policy.list.dependency",
                    blocked
                        ? "The List dependency has no executable materialization plan."
                        : "Materialize or reuse the captured List dependency closure.",
                    listPlan.TargetRootFolderServerRelativeUrl,
                    "The target List identity and portable schema match the sealed List plan."));

                PublishingPageListSchemaIngredientActionProjector.ProjectList(
                    source,
                    listPlan,
                    blocked,
                    actions,
                    transactionDependencyProjection);
                PublishingPageListContentIngredientActionProjector.Project(
                    source,
                    listPlan,
                    blocked,
                    actions,
                    transactionDependencyProjection);
                AddPlatformFeatures(source.SourceSiteId, listPlan, actions);
                AddViewRenderingResources(source.SourceSiteId, listPlan, actions);
                AddViews(
                    source.SourceWebId,
                    source.SourceListId,
                    listPlan,
                    blocked,
                    actions,
                    transactionDependencyProjection);
            }
        }

        private static bool IsListObjectBlocked(
            PnP.Framework.Migration.Lists.Capture.ListDependencySnapshot source,
            ListMaterializationPlan plan)
        {
            if (source == null
                || source.Availability == EvidenceAvailability.Unavailable
                || source.Availability == EvidenceAvailability.Conflict
                || plan == null)
            {
                return true;
            }

            return (plan.Issues ?? Array.Empty<MigrationIssue>()).Any(value => value != null
                && (value.Severity == MigrationIssueSeverity.Blocker
                    || value.Severity == MigrationIssueSeverity.Error)
                && ListObjectIssueCodes.Contains(value.Code));
        }

        private static void AddViewRenderingResources(
            Guid sourceSiteId,
            ListMaterializationPlan listPlan,
            IDictionary<string, PageIngredientAction> actions)
        {
            foreach (var resource in listPlan.ViewRenderingResources)
            {
                var blocked = resource.Disposition == ListViewRenderingResourceMaterializationDisposition.Block;
                var referenceOnly = resource.Disposition == ListViewRenderingResourceMaterializationDisposition.PreserveReferenceOnly;
                PublishingPageIngredientActionFactory.Add(actions, PublishingPageIngredientActionFactory.Create(
                    PublishingPageIngredientIds.ViewRenderingResource(sourceSiteId, resource.SourceResourceId),
                    blocked ? IngredientCapability.Incompatible : IngredientCapability.Available,
                    blocked
                        ? IngredientDisposition.Block
                        : referenceOnly ? IngredientDisposition.Substitute : IngredientDisposition.Preserve,
                    blocked
                        ? "none"
                        : referenceOnly ? "preserve-mapped-reference-without-resource-bytes" : "copy-exact-bytes-create-only",
                    "policy.list-view.rendering-resource",
                    resource.Reason,
                    blocked ? null : resource.TargetServerRelativeUrl,
                    blocked
                        ? null
                        : referenceOnly
                            ? "No target resource bytes are created; fresh View readback retains the captured reference."
                            : "Fresh resource readback matches the sealed SHA-256."));
            }
        }

        private static void AddPlatformFeatures(
            Guid sourceSiteId,
            ListMaterializationPlan listPlan,
            IDictionary<string, PageIngredientAction> actions)
        {
            foreach (var feature in listPlan.RequiredFeatures)
            {
                // Platform capabilities are evaluated independently from the List that consumes
                // them. A List ownership/schema gap must not misclassify an activatable feature
                // as incompatible; the dependency edge keeps the List transaction gated.
                var blocked = !feature.IsExecutable;
                var isActive = feature.TargetProbe != null && feature.TargetProbe.IsActive;
                PublishingPageIngredientActionFactory.Add(actions, PublishingPageIngredientActionFactory.Create(
                    PublishingPageIngredientIds.PlatformFeature(sourceSiteId, feature.FeatureId),
                    blocked ? IngredientCapability.Incompatible : IngredientCapability.Available,
                    blocked ? IngredientDisposition.Block : IngredientDisposition.Substitute,
                    blocked
                        ? "none"
                        : isActive ? "reuse-target-runtime-feature" : "activate-target-runtime-feature",
                    "policy.platform-feature.required-runtime",
                    blocked
                        ? "The required SharePoint platform feature has no admitted target action."
                        : feature.Reason,
                    blocked ? null : feature.TargetWebUrl + "#site-feature:" + feature.FeatureId.ToString("D"),
                    blocked ? null : "Fresh readback verifies that the target feature is active.",
                    blocked ? null : "Every runtime content type promised by the feature is available before List content-type membership is applied."));
            }
        }

        private static void AddViews(
            Guid sourceWebId,
            Guid sourceListId,
            ListMaterializationPlan listPlan,
            bool listBlocked,
            IDictionary<string, PageIngredientAction> actions,
            bool transactionDependencyProjection)
        {
            foreach (var viewPlan in listPlan.Views)
            {
                var disposition = (!transactionDependencyProjection && listBlocked)
                    || viewPlan.Disposition == ListViewMaterializationDisposition.Block
                    ? IngredientDisposition.Block
                    : viewPlan.Disposition == ListViewMaterializationDisposition.SkipPersonal
                        ? IngredientDisposition.Drop
                        : IngredientDisposition.Preserve;
                PublishingPageIngredientActionFactory.Add(actions, PublishingPageIngredientActionFactory.Create(
                    PublishingPageIngredientIds.View(sourceWebId, sourceListId, viewPlan.SourceViewId),
                    disposition == IngredientDisposition.Block ? IngredientCapability.Incompatible : IngredientCapability.Available,
                    disposition,
                    disposition == IngredientDisposition.Drop
                        ? "omit-personal-view"
                        : disposition == IngredientDisposition.Block ? "none" : "create-or-reuse-view",
                    "policy.list-view.dependency",
                    viewPlan.Reason,
                    listPlan.TargetRootFolderServerRelativeUrl + "#view:" + viewPlan.SourceViewId.ToString("D"),
                    disposition == IngredientDisposition.Drop
                        ? "The omitted personal View remains represented in the source snapshot and reviewed action."
                        : "The target View identity and schema match the sealed View plan."));
            }
        }
    }
}
