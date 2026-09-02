using PnP.Framework.Migration.Pages.Ingredients;
using PnP.Framework.Migration.Pages.Publishing.Capture;
using PnP.Framework.Migration.Pages.Publishing.Planning;
using System;
using System.Collections.Generic;
using System.Linq;

namespace PnP.Framework.Migration.Pages.Publishing.Ingredients
{
    internal static class PublishingPageIngredientActionProjector
    {
        public static IList<PageIngredientAction> Project(
            PublishingPageCaptureBundle snapshot,
            PublishingPageMigrationPlan plan)
        {
            var actions = new Dictionary<string, PageIngredientAction>(StringComparer.Ordinal);
            PublishingPageCoreIngredientActionProjector.Project(snapshot, plan, actions);
            PublishingPageLayoutIngredientActionProjector.Project(snapshot, plan, actions);
            PublishingPageTopologyIngredientActionProjector.Project(snapshot, plan, actions);
            PublishingPageWebPartIngredientActionProjector.Project(plan, actions);
            PublishingPageListIngredientActionProjector.Project(snapshot, plan, actions);
            PublishingPageReferenceIngredientActionProjector.Project(plan, actions);

            foreach (var node in snapshot.IngredientGraph.Nodes.Where(value => value != null && value.HasContent))
            {
                if (!actions.ContainsKey(node.Id))
                {
                    PublishingPageIngredientActionFactory.Add(actions, PublishingPageIngredientActionFactory.Create(
                        node.Id,
                        IngredientCapability.Unknown,
                        IngredientDisposition.Block,
                        "none",
                        "policy.ingredient.unknown",
                        "No ingredient handler produced an action for the captured ingredient."));
                }
            }

            return actions.Values.OrderBy(value => value.IngredientId, StringComparer.Ordinal).ToList();
        }
    }
}
