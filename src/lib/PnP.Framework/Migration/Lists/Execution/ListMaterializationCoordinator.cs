using Microsoft.SharePoint.Client;
using PnP.Framework.Migration.Execution;
using PnP.Framework.Migration.Lists.Capture;
using PnP.Framework.Migration.Lists.Planning;
using PnP.Framework.Migration.Packaging;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using PnP.Framework.Migration.Schema.ContentTypes.Execution;
using PnP.Framework.Migration.Schema.Fields;
using PnP.Framework.Migration.Features;

namespace PnP.Framework.Migration.Lists.Execution
{
    internal static class ListMaterializationCoordinator
    {
        public static IDictionary<Guid, ListMaterializationReceipt> Ensure(
            ClientContext anchorContext,
            IEnumerable<ListDependencySnapshot> snapshots,
            ListMigrationPlanSet planSet,
            MigrationExecutionRecorder recorder,
            IMigrationArtifactStore artifactStore)
        {
            return Ensure(anchorContext, snapshots, planSet, null, recorder, artifactStore);
        }

        public static IDictionary<Guid, ListMaterializationReceipt> Ensure(
            ClientContext anchorContext,
            IEnumerable<ListDependencySnapshot> snapshots,
            ListMigrationPlanSet planSet,
            ListMaterializationExecutionScope executionScope,
            MigrationExecutionRecorder recorder,
            IMigrationArtifactStore artifactStore)
        {
            var capturedSources = (snapshots ?? Enumerable.Empty<ListDependencySnapshot>())
                .Where(value => value != null)
                .ToDictionary(value => value.SourceListId);
            if (executionScope == null && capturedSources.Count == 0)
            {
                recorder.RecordAlreadySatisfied("lists.materialize", "The approved source snapshot has no List dependencies.");
                return new Dictionary<Guid, ListMaterializationReceipt>();
            }
            if (executionScope?.HasWork == false)
            {
                recorder.RecordAlreadySatisfied("lists.materialize", "No List ingredient is present in the admitted execution frontier.");
                return new Dictionary<Guid, ListMaterializationReceipt>();
            }
            if (planSet == null || (executionScope == null && !planSet.IsExecutable))
            {
                throw new InvalidOperationException("A complete executable List migration plan is required before materialization.");
            }

            var capturedPlans = planSet.Lists
                .Where(value => value != null)
                .ToDictionary(value => value.SourceListId);
            if (executionScope == null
                && (capturedPlans.Count != capturedSources.Count
                    || !new HashSet<Guid>(capturedPlans.Keys).SetEquals(capturedSources.Keys)))
            {
                throw new InvalidDataException("The List plan set does not exactly cover the captured List dependency closure.");
            }

            var receipts = new Dictionary<Guid, ListMaterializationReceipt>();
            EnsurePlatformFeatures(anchorContext, planSet, executionScope, recorder);
            EnsureSiteFields(anchorContext, planSet, executionScope, recorder);
            ContentTypeClosureMaterializer.Ensure(
                anchorContext,
                planSet.Lists.SelectMany(value => value.SiteContentTypes)
                    .Where(value => executionScope == null || executionScope.IncludesSiteContentType(value)),
                recorder);
            EnsureStandaloneViewRenderingResources(
                anchorContext,
                capturedPlans,
                executionScope,
                recorder,
                artifactStore);

            var selectedListIds = executionScope == null
                ? new HashSet<Guid>(capturedSources.Keys)
                : new HashSet<Guid>(executionScope.Lists
                    .Where(value => value.HasListScopedWork)
                    .Select(value => value.SourceListId));
            var orderedListIds = planSet.OrderedSourceListIds
                .Where(selectedListIds.Contains)
                .Concat(selectedListIds.Except(planSet.OrderedSourceListIds).OrderBy(value => value))
                .ToArray();
            foreach (var sourceListId in orderedListIds)
            {
                ListDependencySnapshot source;
                ListMaterializationPlan plan;
                if (!capturedSources.TryGetValue(sourceListId, out var capturedSource)
                    || !capturedPlans.TryGetValue(sourceListId, out var capturedPlan))
                {
                    throw new InvalidDataException("The List dependency order references an unknown source List: " + sourceListId.ToString("D") + ".");
                }
                var selection = executionScope?.GetList(sourceListId);
                if (selection != null && !selection.IncludeListObject)
                {
                    throw new InvalidDataException(
                        "A List-scoped transaction entered the execution frontier without its required List object: "
                        + sourceListId.ToString("D") + ".");
                }
                source = executionScope == null
                    ? capturedSource
                    : executionScope.ProjectSource(capturedSource);
                plan = executionScope == null
                    ? capturedPlan
                    : executionScope.ProjectPlan(capturedPlan);

                using (var context = anchorContext.Clone(plan.TargetWebUrl))
                {
                    context.Load(context.Web, value => value.Id, value => value.Url, value => value.ServerRelativeUrl);
                    context.ExecuteQueryRetry();
                    var prefix = "list." + sourceListId.ToString("N");
                    var targetListResult = recorder.Execute(
                        prefix + ".object",
                        "Ensure migration-owned target List '" + plan.TargetRootFolderServerRelativeUrl + "'.",
                        () => ListObjectMaterializer.Ensure(context, source, plan),
                        value => value.Disposition == ListMaterializationDisposition.ReuseOwned
                            ? MutationOutcome.AlreadySatisfied
                            : MutationOutcome.Applied,
                        value => "Target List " + value.List.Id.ToString("D") + " passed fresh ownership preflight and readback as " + value.Disposition + ".");
                    var targetList = targetListResult.List;

                    var contentTypeIds = EnsureContentTypeMembership(
                        context,
                        targetList,
                        source,
                        recorder,
                        prefix);
                    EnsureListFields(context, targetList, plan, receipts, recorder, prefix);
                    EnsureContentTypeFieldLinks(
                        context,
                        targetList,
                        source,
                        plan,
                        contentTypeIds,
                        recorder,
                        prefix);
                    EnsureContentTypeOrder(
                        context,
                        targetList,
                        source,
                        contentTypeIds,
                        selection,
                        recorder,
                        prefix);
                    var itemIds = EnsureItems(
                        context,
                        targetList,
                        source,
                        plan,
                        selection,
                        receipts,
                        contentTypeIds,
                        artifactStore,
                        recorder,
                        prefix);
                    EnsureViewRenderingResources(context, plan, artifactStore, recorder, prefix);
                    var viewIds = EnsureViews(context, targetList, plan, recorder, prefix);

                    context.Load(targetList, value => value.Id, value => value.RootFolder.ServerRelativeUrl);
                    context.ExecuteQueryRetry();
                    var receipt = new ListMaterializationReceipt
                    {
                        SourceWebId = source.SourceWebId,
                        SourceListId = source.SourceListId,
                        TargetWebId = context.Web.Id,
                        TargetListId = targetList.Id,
                        TargetRootFolderServerRelativeUrl = targetList.RootFolder.ServerRelativeUrl,
                        TargetItemIds = itemIds,
                        TargetViewIds = viewIds,
                        TargetContentTypeIds = contentTypeIds,
                        Disposition = targetListResult.Disposition,
                        PlanDigest = plan.PlanDigest
                    };
                    ListMaterializationVerifier.Verify(context, source, plan, receipt, receipts, selection);
                    receipts[sourceListId] = receipt;
                }
            }
            return receipts;
        }

        private static IDictionary<string, string> EnsureContentTypeMembership(
            ClientContext context,
            List targetList,
            ListDependencySnapshot source,
            MigrationExecutionRecorder recorder,
            string prefix)
        {
            if (source.ContentTypes.Count == 0)
            {
                recorder.RecordAlreadySatisfied(prefix + ".content-types.membership",
                    "No List content type membership is present in the admitted execution frontier.");
                return new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            }
            return recorder.Execute(
                prefix + ".content-types.membership",
                "Ensure target List content type membership.",
                () => ListContentTypeMaterializer.EnsureMembership(context, targetList, source),
                values => values.Count == 0 ? MutationOutcome.AlreadySatisfied : MutationOutcome.Applied,
                values => "Resolved " + values.Count + " source-to-target List content type identities.");
        }

        private static void EnsureListFields(
            ClientContext context,
            List targetList,
            ListMaterializationPlan plan,
            IDictionary<Guid, ListMaterializationReceipt> receipts,
            MigrationExecutionRecorder recorder,
            string prefix)
        {
            if (plan.Fields.Count == 0)
            {
                recorder.RecordAlreadySatisfied(prefix + ".fields",
                    "No List field schema is present in the admitted execution frontier.");
                return;
            }
            recorder.Execute(prefix + ".fields", "Ensure approved target List field schema.", () =>
                ListFieldMaterializer.Ensure(context, targetList, plan, receipts));
        }

        private static void EnsureContentTypeFieldLinks(
            ClientContext context,
            List targetList,
            ListDependencySnapshot source,
            ListMaterializationPlan plan,
            IDictionary<string, string> contentTypeIds,
            MigrationExecutionRecorder recorder,
            string prefix)
        {
            if (source.ContentTypes.Count == 0)
            {
                recorder.RecordAlreadySatisfied(prefix + ".content-types.field-links",
                    "No List content type field-link transaction is present in the admitted execution frontier.");
                return;
            }
            recorder.Execute(prefix + ".content-types.field-links", "Apply approved List content type field-link settings.", () =>
                ListContentTypeMaterializer.EnsureFieldLinks(context, targetList, source, plan, contentTypeIds));
        }

        private static void EnsureContentTypeOrder(
            ClientContext context,
            List targetList,
            ListDependencySnapshot source,
            IDictionary<string, string> contentTypeIds,
            ListMaterializationExecutionScope.ListSelection selection,
            MigrationExecutionRecorder recorder,
            string prefix)
        {
            if (selection != null && !selection.ExactContentTypeInventory)
            {
                recorder.RecordAlreadySatisfied(prefix + ".content-types.order",
                    "Content type order is outside the partial execution frontier; the existing target order is preserved.");
                return;
            }
            recorder.Execute(prefix + ".content-types.order", "Apply the captured List content type order.", () =>
                ListContentTypeMaterializer.EnsureOrder(context, targetList, source, contentTypeIds));
        }

        private static IDictionary<int, int> EnsureItems(
            ClientContext context,
            List targetList,
            ListDependencySnapshot source,
            ListMaterializationPlan plan,
            ListMaterializationExecutionScope.ListSelection selection,
            IDictionary<Guid, ListMaterializationReceipt> receipts,
            IDictionary<string, string> contentTypeIds,
            IMigrationArtifactStore artifactStore,
            MigrationExecutionRecorder recorder,
            string prefix)
        {
            if (source.Items.Count == 0)
            {
                recorder.RecordAlreadySatisfied(prefix + ".items",
                    "No List item, document, folder, or attachment is present in the admitted execution frontier.");
                return new Dictionary<int, int>();
            }
            return recorder.Execute(
                prefix + ".items",
                "Replay captured List items, documents, folders, and attachments in the admitted frontier.",
                () => selection == null
                    ? ListItemMaterializer.Ensure(context, targetList, source, plan, receipts, contentTypeIds, artifactStore)
                    : ListItemMaterializer.Ensure(context, targetList, source, plan, selection, receipts, contentTypeIds, artifactStore),
                values => values.Count == 0 ? MutationOutcome.AlreadySatisfied : MutationOutcome.Applied,
                values => "Resolved " + values.Count + " source-to-target List item identities.");
        }

        private static void EnsureViewRenderingResources(
            ClientContext context,
            ListMaterializationPlan plan,
            IMigrationArtifactStore artifactStore,
            MigrationExecutionRecorder recorder,
            string prefix)
        {
            if (plan.ViewRenderingResources.Count == 0)
            {
                recorder.RecordAlreadySatisfied(prefix + ".view-rendering-resources",
                    "No custom View rendering resource is present in the admitted execution frontier.");
                return;
            }
            recorder.Execute(
                prefix + ".view-rendering-resources",
                "Ensure exact custom View rendering resources.",
                () => ListViewRenderingResourceMaterializer.Ensure(context, plan, artifactStore),
                value => value == 0 ? MutationOutcome.AlreadySatisfied : MutationOutcome.Applied,
                value => "Created " + value + " custom View rendering resources; every reused resource matched the sealed SHA-256.");
        }

        private static IDictionary<Guid, Guid> EnsureViews(
            ClientContext context,
            List targetList,
            ListMaterializationPlan plan,
            MigrationExecutionRecorder recorder,
            string prefix)
        {
            if (plan.Views.Count == 0)
            {
                recorder.RecordAlreadySatisfied(prefix + ".views",
                    "No List View is present in the admitted execution frontier.");
                return new Dictionary<Guid, Guid>();
            }
            return recorder.Execute(
                prefix + ".views",
                "Ensure approved target List views.",
                () => ListViewMaterializer.Ensure(context, targetList, plan),
                values => values.Count == 0 ? MutationOutcome.AlreadySatisfied : MutationOutcome.Applied,
                values => "Resolved " + values.Count + " source-to-target View identities.");
        }

        private static void EnsureStandaloneViewRenderingResources(
            ClientContext anchorContext,
            IDictionary<Guid, ListMaterializationPlan> plans,
            ListMaterializationExecutionScope executionScope,
            MigrationExecutionRecorder recorder,
            IMigrationArtifactStore artifactStore)
        {
            if (executionScope == null)
            {
                return;
            }
            foreach (var selection in executionScope.Lists.Where(value =>
                         !value.HasListScopedWork && value.ViewRenderingResourceIds.Count > 0))
            {
                if (!plans.TryGetValue(selection.SourceListId, out var capturedPlan))
                {
                    throw new InvalidDataException(
                        "A View rendering-resource transaction has no sealed List plan: "
                        + selection.SourceListId.ToString("D") + ".");
                }
                var plan = executionScope.ProjectPlan(capturedPlan);
                using (var context = anchorContext.Clone(plan.TargetWebUrl))
                {
                    var prefix = "list." + selection.SourceListId.ToString("N");
                    EnsureViewRenderingResources(context, plan, artifactStore, recorder, prefix);
                    var diagnostics = new List<string>();
                    var verified = ListViewRenderingResourceMaterializer.Verify(context, plan, diagnostics);
                    if (diagnostics.Count > 0 || verified != plan.ViewRenderingResources.Count)
                    {
                        throw new InvalidOperationException(
                            "Fresh standalone View rendering-resource verification failed: "
                            + string.Join("; ", diagnostics));
                    }
                }
            }
        }

        private static void EnsureSiteFields(
            ClientContext anchorContext,
            ListMigrationPlanSet planSet,
            ListMaterializationExecutionScope executionScope,
            MigrationExecutionRecorder recorder)
        {
            if (executionScope == null)
            {
                return;
            }
            var fields = planSet.Lists
                .SelectMany(value => value.SiteContentTypes)
                .Where(value => value?.Schema != null)
                .SelectMany(node => node.Schema.Fields
                    .Where(field => field != null
                        && executionScope.IncludesSiteField(node.SourceOwnerWebUrl, field.FieldId))
                    .Select(field => new { Node = node, Field = field }))
                .GroupBy(value => value.Node.TargetOwnerWebUrl, StringComparer.OrdinalIgnoreCase)
                .OrderBy(value => value.Key, StringComparer.OrdinalIgnoreCase);
            foreach (var owner in fields)
            {
                using (var context = anchorContext.Clone(owner.Key))
                {
                    var selected = owner.Select(value => value.Field).ToArray();
                    recorder.Execute(
                        "schema.site-fields." + MigrationDigest.ComputeSha256(owner.Key).Substring(0, 16),
                        "Ensure " + selected.Length + " independently executable site field(s) at '" + owner.Key + "'.",
                        () => SiteFieldMaterializer.Ensure(context, context.Web, selected),
                        count => count == 0 ? MutationOutcome.AlreadySatisfied : MutationOutcome.Applied,
                        count => "Created " + count + " site field(s); every selected field passed fresh schema readback.");
                }
            }
        }

        private static void EnsurePlatformFeatures(
            ClientContext anchorContext,
            ListMigrationPlanSet planSet,
            ListMaterializationExecutionScope executionScope,
            MigrationExecutionRecorder recorder)
        {
            var requirements = planSet.Lists
                .SelectMany(list => list.RequiredFeatures.Select(feature => new
                {
                    list.SourceSiteId,
                    Feature = feature
                }))
                .Where(value => executionScope == null
                    || executionScope.IncludesPlatformFeature(value.SourceSiteId, value.Feature.FeatureId))
                .GroupBy(value => new { value.SourceSiteId, value.Feature.FeatureId })
                .Select(group => Merge(group.Key.SourceSiteId, group.Select(value => value.Feature)))
                .OrderBy(value => value.Plan.DependencyOrder)
                .ThenBy(value => value.SourceSiteId)
                .ThenBy(value => value.Plan.FeatureId)
                .ToArray();
            foreach (var requirement in requirements)
            {
                var actionId = "platform-feature.site." + requirement.SourceSiteId.ToString("N") + "."
                    + requirement.Plan.FeatureId.ToString("N");
                recorder.Execute(
                    actionId,
                    "Ensure target site feature '" + requirement.Plan.Name + "' (" + requirement.Plan.FeatureId.ToString("D") + ").",
                    () => PlatformFeatureMaterializer.Ensure(anchorContext, requirement.Plan),
                    activated => activated ? MutationOutcome.Applied : MutationOutcome.AlreadySatisfied,
                    activated => activated
                        ? "Activated and freshly verified the SharePoint-owned site feature."
                        : "The SharePoint-owned site feature and its runtime contract were already satisfied.");
            }
        }

        private static FeatureExecutionRequirement Merge(
            Guid sourceSiteId,
            IEnumerable<PlatformFeatureMaterializationPlan> plans)
        {
            var values = plans.ToArray();
            var first = values.OrderBy(value => new Uri(value.TargetWebUrl).AbsolutePath.Length)
                .ThenBy(value => value.TargetWebUrl, StringComparer.OrdinalIgnoreCase).First();
            if (values.Any(value => value.Scope != first.Scope
                || value.DependencyOrder != first.DependencyOrder
                || !value.DependsOnFeatureIds.SequenceEqual(first.DependsOnFeatureIds)
                || !string.Equals(value.Name, first.Name, StringComparison.Ordinal)
                || !string.Equals(value.TargetWebUrl, first.TargetWebUrl, StringComparison.OrdinalIgnoreCase)
                || !string.Equals(value.Reason, first.Reason, StringComparison.Ordinal)))
            {
                throw new InvalidDataException("Duplicate platform feature requirements disagree for feature " + first.FeatureId.ToString("D") + ".");
            }
            return new FeatureExecutionRequirement
            {
                SourceSiteId = sourceSiteId,
                Plan = new PlatformFeatureMaterializationPlan
                {
                    FeatureId = first.FeatureId,
                    Name = first.Name,
                    Scope = first.Scope,
                    DependencyOrder = first.DependencyOrder,
                    DependsOnFeatureIds = values.SelectMany(value => value.DependsOnFeatureIds).Distinct().OrderBy(value => value).ToList(),
                    RequiredByContentTypeIds = values.SelectMany(value => value.RequiredByContentTypeIds)
                        .Distinct(StringComparer.OrdinalIgnoreCase).OrderBy(value => value, StringComparer.OrdinalIgnoreCase).ToList(),
                    ExpectedContentTypeIds = values.SelectMany(value => value.ExpectedContentTypeIds)
                        .Distinct(StringComparer.OrdinalIgnoreCase).OrderBy(value => value, StringComparer.OrdinalIgnoreCase).ToList(),
                    TargetWebUrl = first.TargetWebUrl,
                    Disposition = values.Any(value => value.Disposition == PlatformFeatureMaterializationDisposition.Block)
                        ? PlatformFeatureMaterializationDisposition.Block
                        : PlatformFeatureMaterializationDisposition.EnsureActive,
                    Reason = first.Reason
                }
            };
        }

        private sealed class FeatureExecutionRequirement
        {
            public Guid SourceSiteId { get; set; }

            public PlatformFeatureMaterializationPlan Plan { get; set; }
        }
    }
}
