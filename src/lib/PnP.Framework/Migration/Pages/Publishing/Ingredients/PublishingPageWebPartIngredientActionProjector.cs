using PnP.Framework.Migration.Pages.ClassicWebParts.Bindings;
using PnP.Framework.Migration.Pages.Ingredients;
using PnP.Framework.Migration.Pages.Publishing.Planning;
using System.Collections.Generic;

namespace PnP.Framework.Migration.Pages.Publishing.Ingredients
{
    internal static class PublishingPageWebPartIngredientActionProjector
    {
        public static void Project(
            PublishingPageMigrationPlan plan,
            IDictionary<string, PageIngredientAction> actions)
        {
            foreach (var webPartAction in plan.WebPartActions)
            {
                var blocked = webPartAction.Disposition == ClassicWebPartDisposition.Block;
                PublishingPageIngredientActionFactory.Add(actions, PublishingPageIngredientActionFactory.Create(
                    PublishingPageIngredientIds.WebPart(webPartAction.SourceWebPartId),
                    blocked ? IngredientCapability.Incompatible : IngredientCapability.Available,
                    blocked ? IngredientDisposition.Block : IngredientDisposition.Preserve,
                    webPartAction.Disposition == ClassicWebPartDisposition.RebindListAfterMaterialization
                        ? "rewrite-list-and-view-binding"
                        : blocked ? "none" : "copy-captured-export",
                    "policy.webpart.classic",
                    webPartAction.Reason,
                    webPartAction.TargetListServerRelativeUrl ?? plan.TargetPageServerRelativeUrl,
                    $"The target shared Web Part corresponding to source '{webPartAction.SourceWebPartId:D}' has approved zone placement and bindings."));
            }
        }
    }
}
