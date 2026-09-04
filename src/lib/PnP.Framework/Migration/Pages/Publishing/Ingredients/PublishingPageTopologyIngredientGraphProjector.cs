using PnP.Framework.Migration.Pages.Ingredients;
using PnP.Framework.Migration.Pages.Publishing.Capture;
using PnP.Framework.Migration.Topology;
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
            Project(snapshot, graph, PublishingPageIngredientGraphProjectionRevision.CurrentV7);
        }

        internal static void Project(
            PublishingPageCaptureBundle snapshot,
            CanonicalPageIngredientGraph graph,
            PublishingPageIngredientGraphProjectionRevision revision)
        {
            var sourceWebs = (snapshot.SourceTopology?.Webs ?? Array.Empty<SourceWebSnapshot>())
                .Where(value => value != null)
                .ToList();
            var webIds = new HashSet<Guid>(sourceWebs.Select(value => value.WebId));
            foreach (var web in sourceWebs.OrderBy(value => value.ServerRelativeUrl, StringComparer.OrdinalIgnoreCase))
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

            // Older or authorization-limited captures can retain the exact page-Web
            // identity while lacking its parent-Web closure. The page Web remains a
            // real content-bearing ingredient: omitting it would collapse a scoped
            // topology authorization failure into an untyped whole-page stop.
            if (snapshot.Source != null
                && snapshot.Source.SiteId != Guid.Empty
                && snapshot.Source.WebId != Guid.Empty
                && revision == PublishingPageIngredientGraphProjectionRevision.CurrentV7
                && !webIds.Contains(snapshot.Source.WebId))
            {
                webIds.Add(snapshot.Source.WebId);
                graph.Nodes.Add(Node(
                    PublishingPageIngredientIds.Web(snapshot.Source.SiteId, snapshot.Source.WebId),
                    PageIngredientKind.Web,
                    snapshot.Source.WebUrl,
                    true,
                    PageIngredientOwnership.Shared,
                    "Captured source page-Web identity; ancestor topology closure was unavailable",
                    null,
                    null));
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
