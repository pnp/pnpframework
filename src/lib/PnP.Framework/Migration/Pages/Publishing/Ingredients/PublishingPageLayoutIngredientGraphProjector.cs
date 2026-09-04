using PnP.Framework.Migration.Pages.Ingredients;
using PnP.Framework.Migration.Pages.Publishing.Capture;
using PnP.Framework.Migration.Pages.Publishing.Layouts;
using PnP.Framework.Migration.Schema.Fields;
using System;
using System.Linq;
using static PnP.Framework.Migration.Pages.Publishing.Ingredients.PublishingPageIngredientGraphFactory;

namespace PnP.Framework.Migration.Pages.Publishing.Ingredients
{
    internal static class PublishingPageLayoutIngredientGraphProjector
    {
        public static void Project(
            PublishingPageCaptureBundle snapshot,
            CanonicalPageIngredientGraph graph,
            PublishingPageIngredientGraphProjectionRevision revision)
        {
            var ownerWebId = PublishingPageIngredientGraphProjector.UsesOwnerWebDependencies(revision)
                ? PublishingPageIngredientOwnerWebResolver.Root(snapshot)
                : null;
            AddContentTypeFields(snapshot, graph, ownerWebId);
            AddResources(snapshot, graph, ownerWebId);
            if (PublishingPageIngredientGraphProjector.UsesTransactionDependencies(revision))
            {
                graph.Edges.Add(Edge(
                    PublishingPageIngredientIds.Layout,
                    PublishingPageIngredientIds.ContentType,
                    PageIngredientRelationship.TypedBy,
                    PageIngredientRequirement.Required));
            }
            if (!string.IsNullOrWhiteSpace(ownerWebId))
            {
                graph.Edges.Add(Edge(
                    PublishingPageIngredientIds.Layout,
                    ownerWebId,
                    PageIngredientRelationship.DependsOn,
                    PageIngredientRequirement.Required));
                graph.Edges.Add(Edge(
                    PublishingPageIngredientIds.ContentType,
                    ownerWebId,
                    PageIngredientRelationship.DependsOn,
                    PageIngredientRequirement.Required));
            }
        }

        private static void AddContentTypeFields(
            PublishingPageCaptureBundle snapshot,
            CanonicalPageIngredientGraph graph,
            string ownerWebId)
        {
            foreach (var field in (snapshot.Layout?.AssociatedContentTypeSchema?.RequiredFieldClosure
                         ?? Array.Empty<FieldSchemaSnapshot>())
                     .Where(value => value != null)
                     .GroupBy(value => value.Id)
                     .Select(group => group.First())
                     .OrderBy(value => value.Id))
            {
                var id = PublishingPageIngredientIds.PageContentTypeField(field.Id);
                graph.Nodes.Add(Node(
                    id,
                    PageIngredientKind.Field,
                    field.InternalName,
                    true,
                    PageIngredientOwnership.Shared,
                    "Associated Publishing Content Type field-schema closure",
                    field.PortableSchemaSha256,
                    snapshot.Runtime?.AdapterId));
                graph.Edges.Add(Edge(
                    PublishingPageIngredientIds.ContentType,
                    id,
                    PageIngredientRelationship.BindsTo,
                    PageIngredientRequirement.Required));
                if (!string.IsNullOrWhiteSpace(ownerWebId))
                {
                    graph.Edges.Add(Edge(
                        id,
                        ownerWebId,
                        PageIngredientRelationship.DependsOn,
                        PageIngredientRequirement.Required));
                }
            }
        }

        private static void AddResources(
            PublishingPageCaptureBundle snapshot,
            CanonicalPageIngredientGraph graph,
            string ownerWebId)
        {
            foreach (var resourceGroup in (snapshot.Layout?.ResourceArtifacts
                         ?? Array.Empty<PublishingPageLayoutResourceSnapshot>())
                     .Where(value => value != null)
                     .GroupBy(value => value.Reference?.Value ?? value.ResolvedSourceUrl ?? string.Empty, StringComparer.Ordinal)
                     .OrderBy(value => value.Key, StringComparer.Ordinal))
            {
                var resource = resourceGroup.First();
                var id = PublishingPageIngredientIds.LayoutResource(resourceGroup.Key);
                graph.Nodes.Add(Node(
                    id,
                    PageIngredientKind.Asset,
                    resource.ResolvedSourceUrl ?? resource.Reference?.Value,
                    true,
                    PageIngredientOwnership.Shared,
                    "Captured Page Layout rendering resource",
                    resource.Artifact?.Sha256,
                    snapshot.Runtime?.AdapterId));
                graph.Edges.Add(Edge(
                    PublishingPageIngredientIds.Layout,
                    id,
                    PageIngredientRelationship.References,
                    PageIngredientRequirement.Required));
                if (!string.IsNullOrWhiteSpace(ownerWebId))
                {
                    graph.Edges.Add(Edge(
                        id,
                        ownerWebId,
                        PageIngredientRelationship.DependsOn,
                        PageIngredientRequirement.Required));
                }
            }
        }
    }
}
