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
            var sources = (snapshots ?? Enumerable.Empty<ListDependencySnapshot>()).ToDictionary(value => value.SourceListId);
            if (sources.Count == 0)
            {
                recorder.RecordAlreadySatisfied("lists.materialize", "The approved source snapshot has no List dependencies.");
                return new Dictionary<Guid, ListMaterializationReceipt>();
            }
            if (planSet == null || !planSet.IsExecutable)
            {
                throw new InvalidOperationException("A complete executable List migration plan is required before materialization.");
            }

            var plans = planSet.Lists.ToDictionary(value => value.SourceListId);
            if (plans.Count != sources.Count || !new HashSet<Guid>(plans.Keys).SetEquals(sources.Keys))
            {
                throw new InvalidDataException("The List plan set does not exactly cover the captured List dependency closure.");
            }

            var receipts = new Dictionary<Guid, ListMaterializationReceipt>();
            ContentTypeClosureMaterializer.Ensure(
                anchorContext,
                planSet.Lists.SelectMany(value => value.SiteContentTypes),
                recorder);
            foreach (var sourceListId in planSet.OrderedSourceListIds)
            {
                ListDependencySnapshot source;
                ListMaterializationPlan plan;
                if (!sources.TryGetValue(sourceListId, out source) || !plans.TryGetValue(sourceListId, out plan))
                {
                    throw new InvalidDataException("The List dependency order references an unknown source List: " + sourceListId.ToString("D") + ".");
                }

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

                    var contentTypeIds = recorder.Execute(
                        prefix + ".content-types.membership",
                        "Ensure target List content type membership.",
                        () => ListContentTypeMaterializer.EnsureMembership(context, targetList, source),
                        values => values.Count == 0 ? MutationOutcome.AlreadySatisfied : MutationOutcome.Applied,
                        values => "Resolved " + values.Count + " source-to-target List content type identities.");

                    recorder.Execute(prefix + ".fields", "Ensure approved target List field schema.", () =>
                        ListFieldMaterializer.Ensure(context, targetList, plan, receipts));

                    recorder.Execute(prefix + ".content-types.field-links", "Apply approved List content type field-link settings.", () =>
                        ListContentTypeMaterializer.EnsureFieldLinks(context, targetList, source, contentTypeIds));

                    recorder.Execute(prefix + ".content-types.order", "Apply the captured List content type order.", () =>
                        ListContentTypeMaterializer.EnsureOrder(context, targetList, source, contentTypeIds));

                    var itemIds = recorder.Execute(
                        prefix + ".items",
                        "Replay captured List items, documents, folders, and attachments.",
                        () => ListItemMaterializer.Ensure(context, targetList, source, plan, receipts, contentTypeIds, artifactStore),
                        values => values.Count == 0 ? MutationOutcome.AlreadySatisfied : MutationOutcome.Applied,
                        values => "Resolved " + values.Count + " source-to-target List item identities.");

                    var viewIds = recorder.Execute(
                        prefix + ".views",
                        "Ensure approved target List views.",
                        () => ListViewMaterializer.Ensure(context, targetList, plan),
                        values => values.Count == 0 ? MutationOutcome.AlreadySatisfied : MutationOutcome.Applied,
                        values => "Resolved " + values.Count + " source-to-target View identities.");

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
                    ListMaterializationVerifier.Verify(context, source, plan, receipt, receipts);
                    receipts[sourceListId] = receipt;
                }
            }
            return receipts;
        }
    }
}
