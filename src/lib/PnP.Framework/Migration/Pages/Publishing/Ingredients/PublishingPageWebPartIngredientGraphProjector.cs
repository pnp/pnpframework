using PnP.Framework.Migration.Pages.Ingredients;
using PnP.Framework.Migration.Pages.Publishing.Capture;
using System.Linq;
using static PnP.Framework.Migration.Pages.Publishing.Ingredients.PublishingPageIngredientGraphFactory;

namespace PnP.Framework.Migration.Pages.Publishing.Ingredients
{
    internal static class PublishingPageWebPartIngredientGraphProjector
    {
        public static void Project(
            PublishingPageCaptureBundle snapshot,
            CanonicalPageIngredientGraph graph,
            PublishingPageIngredientGraphProjectionRevision revision)
        {
            foreach (var webPart in snapshot.WebParts.OrderBy(value => value.Id))
            {
                var id = PublishingPageIngredientIds.WebPart(webPart.Id);
                graph.Nodes.Add(Node(
                    id,
                    PageIngredientKind.WebPart,
                    webPart.TypeName ?? webPart.Title,
                    true,
                    PageIngredientOwnership.SourceOwned,
                    "Shared Web Part store export",
                    webPart.ExportSha256,
                    webPart.TypeName));
                graph.Edges.Add(PublishingPageIngredientGraphProjector.UsesTransactionDependencies(revision)
                    ? Edge(id, PublishingPageIngredientIds.PageArtifact, PageIngredientRelationship.PlacedIn, PageIngredientRequirement.Required)
                    : Edge(PublishingPageIngredientIds.PageArtifact, id, PageIngredientRelationship.PlacedIn, PageIngredientRequirement.Optional));
            }

            foreach (var binding in snapshot.ListWebPartBindings.Where(value => value != null))
            {
                var webPartId = PublishingPageIngredientIds.WebPart(binding.SourceWebPartId);
                var listId = PublishingPageIngredientIds.List(binding.SourceListWebId, binding.SourceListId);
                graph.Edges.Add(Edge(webPartId, listId, PageIngredientRelationship.BindsTo, PageIngredientRequirement.Required));
                if (binding.SourceViewId.HasValue)
                {
                    graph.Edges.Add(Edge(
                        webPartId,
                        PublishingPageIngredientIds.View(binding.SourceListWebId, binding.SourceListId, binding.SourceViewId.Value),
                        PageIngredientRelationship.BindsTo,
                        PageIngredientRequirement.Required));
                }
            }
        }
    }
}
