using PnP.Framework.Migration.Pages.Ingredients;
using PnP.Framework.Migration.Pages.Publishing.Capture;
using System;

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
            return graph;
        }
    }
}
