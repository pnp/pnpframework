using PnP.Framework.Migration.Pages.Ingredients;
using PnP.Framework.Migration.Pages.Publishing.Capture;
using System;
using System.Linq;

namespace PnP.Framework.Migration.Pages.Publishing.Ingredients
{
    internal static class PublishingPageIngredientGraphProjector
    {
        public static CanonicalPageIngredientGraph Project(PublishingPageCaptureBundle snapshot)
        {
            if (snapshot == null)
            {
                throw new ArgumentNullException(nameof(snapshot));
            }

            var graph = new CanonicalPageIngredientGraph();
            PublishingPageCoreIngredientGraphProjector.Project(snapshot, graph);
            PublishingPageLayoutIngredientGraphProjector.Project(snapshot, graph);
            PublishingPageTopologyIngredientGraphProjector.Project(snapshot, graph);
            PublishingPageWebPartIngredientGraphProjector.Project(snapshot, graph);
            PublishingPageListIngredientGraphProjector.Project(snapshot, graph);
            PublishingPageReferenceIngredientGraphProjector.Project(snapshot, graph);
            // Captured inventories can contain the same semantic binding more than once
            // (for example, duplicate ListItem field-value evidence). Ingredient edges are
            // set-valued, so collapse only edges whose complete semantic identity matches.
            graph.Edges = graph.Edges
                .GroupBy(
                    edge => edge.FromIngredientId + "\u001f" + edge.ToIngredientId + "\u001f"
                        + edge.Relationship + "\u001f" + edge.Requirement + "\u001f" + edge.Condition,
                    StringComparer.Ordinal)
                .Select(group => group.First())
                .ToList();
            return graph;
        }
    }
}
