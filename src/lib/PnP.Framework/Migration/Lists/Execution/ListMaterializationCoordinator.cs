using Microsoft.SharePoint.Client;
using PnP.Framework.Migration.Execution;
using PnP.Framework.Migration.Lists.Capture;
using PnP.Framework.Migration.Lists.Planning;
using PnP.Framework.Migration.Packaging;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

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
                    var targetList = recorder.Execute(
                        prefix + ".object",
                        "Ensure migration-owned target List '" + plan.TargetRootFolderServerRelativeUrl + "'.",
                        () => ListObjectMaterializer.Ensure(context, source, plan),
                        value => plan.Disposition == ListMaterializationDisposition.ReuseOwned
                            ? MutationOutcome.AlreadySatisfied
                            : MutationOutcome.Applied,
                        value => "Target List " + value.Id.ToString("D") + " passed fresh ownership preflight and readback.");

                    recorder.Execute(prefix + ".fields", "Ensure approved target List field schema.", () =>
                        ListFieldMaterializer.Ensure(context, targetList, plan, receipts));

                    var itemIds = recorder.Execute(
                        prefix + ".items",
                        "Replay captured List items, documents, folders, and attachments.",
                        () => ListItemMaterializer.Ensure(context, targetList, source, plan, receipts, artifactStore),
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
                    receipts[sourceListId] = new ListMaterializationReceipt
                    {
                        SourceWebId = source.SourceWebId,
                        SourceListId = source.SourceListId,
                        TargetWebId = context.Web.Id,
                        TargetListId = targetList.Id,
                        TargetRootFolderServerRelativeUrl = targetList.RootFolder.ServerRelativeUrl,
                        TargetItemIds = itemIds,
                        TargetViewIds = viewIds,
                        PlanDigest = plan.PlanDigest
                    };
                }
            }
            return receipts;
        }
    }
}
