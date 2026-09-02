using PnP.Framework.Migration.Pages.Ingredients;
using PnP.Framework.Migration.Pages.Publishing.Capture;
using System;
using System.Linq;
using static PnP.Framework.Migration.Pages.Publishing.Ingredients.PublishingPageIngredientGraphFactory;

namespace PnP.Framework.Migration.Pages.Publishing.Ingredients
{
    internal static class PublishingPageReferenceIngredientGraphProjector
    {
        public static void Project(PublishingPageCaptureBundle snapshot, CanonicalPageIngredientGraph graph)
        {
            foreach (var reference in snapshot.Dependencies.Where(value => value != null).OrderBy(value => value.Id, StringComparer.Ordinal))
            {
                var id = PublishingPageIngredientIds.Reference(reference.Id);
                graph.Nodes.Add(Node(
                    id,
                    PageIngredientKind.Reference,
                    reference.OriginalValue,
                    !string.IsNullOrWhiteSpace(reference.OriginalValue),
                    PageIngredientOwnership.SourceOwned,
                    reference.Consumer,
                    reference.ContentSha256,
                    null));
                graph.Edges.Add(Edge(PublishingPageIngredientIds.PageArtifact, id, PageIngredientRelationship.References, PageIngredientRequirement.Optional));
            }
        }
    }
}
