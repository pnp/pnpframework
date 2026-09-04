using PnP.Framework.Migration.Pages.Ingredients;
using PnP.Framework.Migration.Pages.Publishing.Planning;
using PnP.Framework.Migration.Pages.References;
using System;
using System.Collections.Generic;
using System.Linq;

namespace PnP.Framework.Migration.Pages.Publishing.Ingredients
{
    internal static class PublishingPageReferenceIngredientActionProjector
    {
        public static void Project(
            PublishingPageMigrationPlan plan,
            IDictionary<string, PageIngredientAction> actions,
            CanonicalPageIngredientGraph graph)
        {
            foreach (var referenceAction in plan.DependencyActions)
            {
                var mapping = Map(referenceAction.Disposition);
                var ingredientId = PublishingPageIngredientIds.Reference(referenceAction.SnapshotDependencyId);
                var action = PublishingPageIngredientActionFactory.Create(
                    ingredientId,
                    mapping.Capability,
                    mapping.Disposition,
                    mapping.Realization,
                    "policy.reference.page",
                    string.Join("; ", referenceAction.Diagnostics ?? new List<string>()),
                    referenceAction.TargetServerRelativeUrl ?? referenceAction.TargetAbsoluteUrl,
                    $"The reference disposition '{referenceAction.Disposition}' is reflected in stored content and target dependency evidence.");
                if (referenceAction.Disposition == PageReferenceDisposition.MaterializeAtTarget
                    && (graph?.Edges ?? Array.Empty<PageIngredientEdge>()).Any(value => value != null
                        && value.Requirement == PageIngredientRequirement.Required
                        && string.Equals(value.FromIngredientId, ingredientId, StringComparison.Ordinal)
                        && string.Equals(value.ToIngredientId, PublishingPageIngredientIds.PublishingContent, StringComparison.Ordinal)))
                {
                    // Asset copying is independently executable. It does not require the
                    // page-content transaction that will eventually consume its target URL.
                    action.Disposition = IngredientDisposition.Transform;
                    action.ReleasedDependencyIngredientIds.Add(PublishingPageIngredientIds.PublishingContent);
                }
                PublishingPageIngredientActionFactory.Add(actions, action);
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
