using PnP.Framework.Migration.Lists.Capture;
using PnP.Framework.Migration.Pages.ClassicWebParts.Bindings;
using PnP.Framework.Migration.Pages.Ingredients;
using PnP.Framework.Migration.Pages.Publishing.Capture;
using PnP.Framework.Migration.Schema.ContentTypes;
using System;
using System.Collections.Generic;
using System.Linq;
using static PnP.Framework.Migration.Pages.Publishing.Ingredients.PublishingPageIngredientGraphFactory;

namespace PnP.Framework.Migration.Pages.Publishing.Ingredients
{
    internal static class PublishingPageListIngredientGraphProjector
    {
        public static void Project(
            PublishingPageCaptureBundle snapshot,
            CanonicalPageIngredientGraph graph,
            PublishingPageIngredientGraphProjectionRevision revision)
        {
            PublishingPageListSchemaIngredientGraphProjector.ProjectSharedClosures(snapshot, graph, revision);
            var lists = (snapshot.ListDependencies ?? Array.Empty<ListDependencySnapshot>())
                .Where(value => value != null)
                .ToArray();
            var listsById = lists.ToDictionary(value => value.SourceListId);
            ProjectMissingBindingPlaceholders(snapshot, lists, graph);
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

                PublishingPageListSchemaIngredientGraphProjector.ProjectList(list, listId, listsById, graph, revision);
                PublishingPageListContentIngredientGraphProjector.Project(
                    list,
                    listId,
                    graph,
                    revision,
                    PublishingPageIngredientOwnerWebResolver.Root(snapshot));
                if (revision != PublishingPageIngredientGraphProjectionRevision.LegacyV1)
                {
                    ProjectPlatformFeatures(snapshot, list, listId, graph, revision);
                }
            }
            AddLookupEdges(snapshot, graph);
        }

        private static void ProjectMissingBindingPlaceholders(
            PublishingPageCaptureBundle snapshot,
            ListDependencySnapshot[] capturedLists,
            CanonicalPageIngredientGraph graph)
        {
            var captured = new HashSet<string>(
                capturedLists.Select(value => ListKey(value.SourceWebId, value.SourceListId)),
                StringComparer.OrdinalIgnoreCase);
            foreach (var bindingGroup in (snapshot.ListWebPartBindings ?? Array.Empty<ClassicListWebPartBindingSnapshot>())
                         .Where(value => value != null)
                         .Where(value => !captured.Contains(ListKey(value.SourceListWebId, value.SourceListId)))
                         .GroupBy(value => ListKey(value.SourceListWebId, value.SourceListId), StringComparer.OrdinalIgnoreCase))
            {
                var binding = bindingGroup.First();
                var listId = PublishingPageIngredientIds.List(binding.SourceListWebId, binding.SourceListId);
                graph.Nodes.Add(Node(
                    listId,
                    PageIngredientKind.List,
                    "Uncaptured List " + binding.SourceListId.ToString("D"),
                    false,
                    PageIngredientOwnership.Shared,
                    "Classic List Web Part binding retained after List closure capture failed",
                    null,
                    "A complete source List snapshot is required before this dependency can be materialized."));
                if (snapshot.SourceTopology?.Webs.Any(value => value.WebId == binding.SourceListWebId) == true)
                {
                    graph.Edges.Add(Edge(
                        listId,
                        PublishingPageIngredientIds.Web(snapshot.Source.SiteId, binding.SourceListWebId),
                        PageIngredientRelationship.DependsOn,
                        PageIngredientRequirement.Required));
                }

                foreach (var viewId in bindingGroup
                             .Where(value => value.SourceViewId.HasValue)
                             .Select(value => value.SourceViewId.Value)
                             .Distinct()
                             .OrderBy(value => value))
                {
                    graph.Nodes.Add(Node(
                        PublishingPageIngredientIds.View(binding.SourceListWebId, binding.SourceListId, viewId),
                        PageIngredientKind.View,
                        "Uncaptured View " + viewId.ToString("D"),
                        false,
                        PageIngredientOwnership.Shared,
                        "Classic List Web Part binding retained after List/View closure capture failed",
                        null,
                        "A complete source View snapshot is required before this binding can be materialized."));
                }
            }
        }

        private static string ListKey(Guid webId, Guid listId)
        {
            return webId.ToString("D") + ":" + listId.ToString("D");
        }

        private static void ProjectPlatformFeatures(
            PublishingPageCaptureBundle snapshot,
            ListDependencySnapshot list,
            string listId,
            CanonicalPageIngredientGraph graph,
            PublishingPageIngredientGraphProjectionRevision revision)
        {
            var ownerWebId = PublishingPageIngredientGraphProjector.UsesOwnerWebDependencies(revision)
                ? PublishingPageIngredientOwnerWebResolver.Root(snapshot)
                : null;
            var features = ContentTypeRuntimeCatalog.CreateFeatureRequirements(
                list.ContentTypes.Select(value => value.ParentId),
                list.SiteContentTypes,
                list.SourceWebUrl);
            foreach (var feature in features)
            {
                var featureId = PublishingPageIngredientIds.PlatformFeature(list.SourceSiteId, feature.FeatureId);
                if (!graph.Nodes.Any(value => string.Equals(value.Id, featureId, StringComparison.Ordinal)))
                {
                    graph.Nodes.Add(Node(
                        featureId,
                        PageIngredientKind.PlatformFeature,
                        feature.Name,
                        true,
                        PageIngredientOwnership.TargetRuntime,
                        "Captured List content-type parent relationship plus the SharePoint runtime capability catalog",
                        null,
                        "The mapped target site collection must expose the platform feature and its expected runtime content types."));
                }
                graph.Edges.Add(Edge(listId, featureId, PageIngredientRelationship.DependsOn, PageIngredientRequirement.Required));
                if (!string.IsNullOrWhiteSpace(ownerWebId))
                {
                    graph.Edges.Add(Edge(
                        featureId,
                        ownerWebId,
                        PageIngredientRelationship.DependsOn,
                        PageIngredientRequirement.Required));
                }
                foreach (var dependencyId in feature.DependsOnFeatureIds)
                {
                    graph.Edges.Add(Edge(
                        featureId,
                        PublishingPageIngredientIds.PlatformFeature(list.SourceSiteId, dependencyId),
                        PageIngredientRelationship.DependsOn,
                        PageIngredientRequirement.Required));
                }
            }
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
