using PnP.Framework.Migration.Lists.Capture;
using PnP.Framework.Migration.Pages.Ingredients;
using PnP.Framework.Migration.Pages.Publishing.Capture;
using System;
using System.Linq;
using static PnP.Framework.Migration.Pages.Publishing.Ingredients.PublishingPageIngredientGraphFactory;

namespace PnP.Framework.Migration.Pages.Publishing.Ingredients
{
    internal static class PublishingPageListIngredientGraphProjector
    {
        public static void Project(PublishingPageCaptureBundle snapshot, CanonicalPageIngredientGraph graph)
        {
            PublishingPageListSchemaIngredientGraphProjector.ProjectSharedClosures(snapshot, graph);
            var lists = (snapshot.ListDependencies ?? Array.Empty<ListDependencySnapshot>())
                .Where(value => value != null)
                .ToArray();
            var listsById = lists.ToDictionary(value => value.SourceListId);
            foreach (var list in lists
                         .OrderBy(value => value.SourceWebId)
                         .ThenBy(value => value.SourceListId))
            {
                var listId = PublishingPageIngredientIds.List(list.SourceWebId, list.SourceListId);
                graph.Nodes.Add(Node(
                    listId,
                    PageIngredientKind.List,
                    list.Title,
                    true,
                    PageIngredientOwnership.Shared,
                    "Captured List dependency closure",
                    null,
                    null));
                if (snapshot.SourceTopology?.Webs.Any(value => value.WebId == list.SourceWebId) == true)
                {
                    graph.Edges.Add(Edge(
                        listId,
                        PublishingPageIngredientIds.Web(list.SourceSiteId, list.SourceWebId),
                        PageIngredientRelationship.DependsOn,
                        PageIngredientRequirement.Required));
                }

                PublishingPageListSchemaIngredientGraphProjector.ProjectList(list, listId, listsById, graph);
                PublishingPageListContentIngredientGraphProjector.Project(list, listId, graph);
            }
            AddLookupEdges(snapshot, graph);
        }

        private static void AddLookupEdges(PublishingPageCaptureBundle snapshot, CanonicalPageIngredientGraph graph)
        {
            var lists = snapshot.ListDependencies.ToDictionary(value => value.SourceListId);
            foreach (var dependency in snapshot.ListLookupDependencies.Where(value => value != null))
            {
                if (lists.TryGetValue(dependency.SourceListId, out var consumer)
                    && lists.TryGetValue(dependency.LookupListId, out var provider))
                {
                    graph.Edges.Add(Edge(
                        PublishingPageIngredientIds.List(consumer.SourceWebId, consumer.SourceListId),
                        PublishingPageIngredientIds.List(provider.SourceWebId, provider.SourceListId),
                        PageIngredientRelationship.DependsOn,
                        PageIngredientRequirement.Required));
                }
            }
        }
    }
}
