using PnP.Framework.Migration.Pages.Ingredients;
using PnP.Framework.Migration.Pages.Publishing.Capture;
using System;
using System.Collections.Generic;
using System.Linq;
using static PnP.Framework.Migration.Pages.Publishing.Ingredients.PublishingPageIngredientGraphFactory;

namespace PnP.Framework.Migration.Pages.Publishing.Ingredients
{
    internal static class PublishingPageTopologyIngredientGraphProjector
    {
        public static void Project(PublishingPageCaptureBundle snapshot, CanonicalPageIngredientGraph graph)
        {
            if (snapshot.SourceTopology == null)
            {
                return;
            }

            var webIds = new HashSet<Guid>(snapshot.SourceTopology.Webs.Select(value => value.WebId));
            foreach (var web in snapshot.SourceTopology.Webs.OrderBy(value => value.ServerRelativeUrl, StringComparer.OrdinalIgnoreCase))
            {
                var id = PublishingPageIngredientIds.Web(web.SiteId, web.WebId);
                graph.Nodes.Add(Node(
                    id,
                    PageIngredientKind.Web,
                    web.WebUrl,
                    true,
                    PageIngredientOwnership.Shared,
                    "Captured source Site/Web ownership closure",
                    null,
                    null));
                if (web.ParentWebId.HasValue && webIds.Contains(web.ParentWebId.Value))
                {
                    graph.Edges.Add(Edge(
                        id,
                        PublishingPageIngredientIds.Web(web.SiteId, web.ParentWebId.Value),
                        PageIngredientRelationship.DependsOn,
                        PageIngredientRequirement.Required));
                }
            }

            if (webIds.Contains(snapshot.Source.WebId))
            {
                graph.Edges.Add(Edge(
                    PublishingPageIngredientIds.PageArtifact,
                    PublishingPageIngredientIds.Web(snapshot.Source.SiteId, snapshot.Source.WebId),
                    PageIngredientRelationship.DependsOn,
                    PageIngredientRequirement.Required));
            }
        }
    }
}
