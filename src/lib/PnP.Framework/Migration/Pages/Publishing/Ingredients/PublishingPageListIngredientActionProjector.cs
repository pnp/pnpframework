using PnP.Framework.Migration.Lists.Planning;
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
        public static void Project(
            PublishingPageCaptureBundle snapshot,
            PublishingPageMigrationPlan plan,
            IDictionary<string, PageIngredientAction> actions)
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

                var blocked = !listPlan.IsExecutable;
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

                PublishingPageListSchemaIngredientActionProjector.ProjectList(source, listPlan, blocked, actions);
                PublishingPageListContentIngredientActionProjector.Project(source, listPlan, blocked, actions);
                AddViews(source.SourceWebId, source.SourceListId, listPlan, blocked, actions);
            }
        }

        private static void AddViews(
            Guid sourceWebId,
            Guid sourceListId,
            ListMaterializationPlan listPlan,
            bool listBlocked,
            IDictionary<string, PageIngredientAction> actions)
        {
            foreach (var viewPlan in listPlan.Views)
            {
                var disposition = listBlocked || viewPlan.Disposition == ListViewMaterializationDisposition.Block
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
