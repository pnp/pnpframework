using PnP.Framework.Migration.Pages.Ingredients;
using PnP.Framework.Migration.Pages.Publishing.Planning;
using PnP.Framework.Migration.Pages.References;
using System.Collections.Generic;

namespace PnP.Framework.Migration.Pages.Publishing.Ingredients
{
    internal static class PublishingPageReferenceIngredientActionProjector
    {
        public static void Project(
            PublishingPageMigrationPlan plan,
            IDictionary<string, PageIngredientAction> actions)
        {
            foreach (var referenceAction in plan.DependencyActions)
            {
                var mapping = Map(referenceAction.Disposition);
                PublishingPageIngredientActionFactory.Add(actions, PublishingPageIngredientActionFactory.Create(
                    PublishingPageIngredientIds.Reference(referenceAction.SnapshotDependencyId),
                    mapping.Capability,
                    mapping.Disposition,
                    mapping.Realization,
                    "policy.reference.page",
                    string.Join("; ", referenceAction.Diagnostics ?? new List<string>()),
                    referenceAction.TargetServerRelativeUrl ?? referenceAction.TargetAbsoluteUrl,
                    $"The reference disposition '{referenceAction.Disposition}' is reflected in stored content and target dependency evidence."));
            }
        }

        private static (IngredientCapability Capability, IngredientDisposition Disposition, string Realization) Map(
            PageReferenceDisposition disposition)
        {
            switch (disposition)
            {
                case PageReferenceDisposition.PreserveExternal:
                    return (IngredientCapability.Available, IngredientDisposition.Preserve, "reuse-external-reference");
                case PageReferenceDisposition.RewriteToTarget:
                    return (IngredientCapability.Available, IngredientDisposition.Transform, "rewrite-reference");
                case PageReferenceDisposition.MaterializeAtTarget:
                    return (IngredientCapability.Available, IngredientDisposition.Preserve, "create-owned");
                case PageReferenceDisposition.Delegate:
                    return (IngredientCapability.Unknown, IngredientDisposition.Delegate, "retain-snapshot");
                default:
                    return (IngredientCapability.Incompatible, IngredientDisposition.Block, "none");
            }
        }
    }
}
