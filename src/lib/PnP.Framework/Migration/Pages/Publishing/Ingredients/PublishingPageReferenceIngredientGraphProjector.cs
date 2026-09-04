using PnP.Framework.Migration.Pages.Ingredients;
using PnP.Framework.Migration.Pages.Capture;
using PnP.Framework.Migration.Pages.Publishing.Capture;
using PnP.Framework.Migration.Pages.References;
using System;
using System.Linq;
using static PnP.Framework.Migration.Pages.Publishing.Ingredients.PublishingPageIngredientGraphFactory;

namespace PnP.Framework.Migration.Pages.Publishing.Ingredients
{
    internal static class PublishingPageReferenceIngredientGraphProjector
    {
        public static void Project(
            PublishingPageCaptureBundle snapshot,
            CanonicalPageIngredientGraph graph,
            PublishingPageIngredientGraphProjectionRevision revision)
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
                graph.Edges.Add(Edge(
                    PublishingPageIngredientGraphProjector.UsesTransactionDependencies(revision)
                        ? PublishingPageIngredientIds.PublishingContent
                        : PublishingPageIngredientIds.PageArtifact,
                    id,
                    PageIngredientRelationship.References,
                    PageIngredientRequirement.Optional));
                if (PublishingPageIngredientGraphProjector.UsesOwnerWebDependencies(revision))
                {
                    graph.Edges.Add(Edge(
                        id,
                        PublishingPageIngredientIds.PublishingContent,
                        PageIngredientRelationship.DependsOn,
                        PageIngredientRequirement.Required));

                    var ownerWebId = MaterializationOwnerWeb(snapshot, reference);
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

        private static string MaterializationOwnerWeb(
            PublishingPageCaptureBundle snapshot,
            PageReferenceSnapshot reference)
        {
            if (snapshot?.Source == null
                || reference == null
                || !reference.IsRenderableResource
                || reference.Kind == PageReferenceKind.IFrame
                || reference.CaptureStatus == PageCaptureStatus.Failed
                || string.IsNullOrWhiteSpace(reference.ContentBase64)
                || string.IsNullOrWhiteSpace(reference.ContentSha256)
                || PageReferenceSnapshotReader.IsSharePointRuntimePath(
                    reference.SourceServerRelativeUrl))
            {
                return null;
            }

            if (!Uri.TryCreate(snapshot.Source.WebUrl, UriKind.Absolute, out var sourceWeb)
                || !Uri.TryCreate(reference.SourceAbsoluteUrl, UriKind.Absolute, out var sourceReference)
                || !string.Equals(sourceWeb.Host, sourceReference.Host, StringComparison.OrdinalIgnoreCase))
            {
                return null;
            }

            return PublishingPageIngredientOwnerWebResolver.ExactOrContaining(
                snapshot,
                reference.SourceServerRelativeUrl ?? reference.SourceAbsoluteUrl);
        }
    }
}
