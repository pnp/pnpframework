using PnP.Framework.Migration.Pages.Fields;
using PnP.Framework.Migration.Pages.Ingredients;
using PnP.Framework.Migration.Pages.Publishing.Capture;
using System;
using System.Collections.Generic;
using System.Linq;
using static PnP.Framework.Migration.Pages.Publishing.Ingredients.PublishingPageIngredientGraphFactory;

namespace PnP.Framework.Migration.Pages.Publishing.Ingredients
{
    internal static class PublishingPageCoreIngredientGraphProjector
    {
        public static void Project(
            PublishingPageCaptureBundle snapshot,
            CanonicalPageIngredientGraph graph,
            PublishingPageIngredientGraphProjectionRevision revision)
        {
            AddRoots(snapshot, graph, revision);
            AddFields(snapshot, graph, revision);
        }

        private static void AddRoots(
            PublishingPageCaptureBundle snapshot,
            CanonicalPageIngredientGraph graph,
            PublishingPageIngredientGraphProjectionRevision revision)
        {
            graph.Nodes.Add(Node(
                PublishingPageIngredientIds.Runtime,
                PageIngredientKind.Runtime,
                snapshot.Runtime?.AdapterId,
                true,
                PageIngredientOwnership.TargetRuntime,
                "ASPX Page directive and layout runtime evidence",
                null,
                snapshot.Runtime?.AdapterId));
            graph.Nodes.Add(Node(
                PublishingPageIngredientIds.PageArtifact,
                PageIngredientKind.PageArtifact,
                snapshot.Source?.PageServerRelativeUrl,
                true,
                PageIngredientOwnership.SourceOwned,
                "Source SPFile bytes",
                snapshot.PageArtifact?.Bytes?.Sha256,
                snapshot.Runtime?.AdapterId));
            graph.Nodes.Add(Node(
                PublishingPageIngredientIds.Layout,
                PageIngredientKind.Layout,
                snapshot.Layout?.FileName,
                true,
                PageIngredientOwnership.Shared,
                "PublishingPageLayout and Page Layout artifact",
                snapshot.Layout?.Bytes?.Sha256,
                snapshot.Runtime?.AdapterId));
            graph.Nodes.Add(Node(
                PublishingPageIngredientIds.ContentType,
                PageIngredientKind.ContentType,
                snapshot.Source?.ContentTypeName,
                !string.IsNullOrWhiteSpace(snapshot.Source?.ContentTypeId),
                PageIngredientOwnership.Shared,
                "Pages library ListItem ContentTypeId",
                null,
                snapshot.Runtime?.AdapterId));
            graph.Nodes.Add(Node(
                PublishingPageIngredientIds.PublishingContent,
                PageIngredientKind.Content,
                "PublishingPageContent",
                !string.IsNullOrEmpty(snapshot.PublishingPageContent),
                PageIngredientOwnership.SourceOwned,
                "Pages library PublishingPageContent field",
                snapshot.PublishingPageContentSha256,
                snapshot.Runtime?.AdapterId));
            graph.Nodes.Add(Node(
                PublishingPageIngredientIds.Security,
                PageIngredientKind.Security,
                "Page security",
                snapshot.Security != null,
                PageIngredientOwnership.Shared,
                "Pages library item role assignments",
                null,
                snapshot.Runtime?.AdapterId));
            graph.Nodes.Add(Node(
                PublishingPageIngredientIds.Lifecycle,
                PageIngredientKind.Lifecycle,
                "Page lifecycle",
                snapshot.Lifecycle != null,
                PageIngredientOwnership.SourceOwned,
                "SPFile and Pages item lifecycle state",
                null,
                snapshot.Runtime?.AdapterId));

            graph.Edges.Add(Edge(PublishingPageIngredientIds.PageArtifact, PublishingPageIngredientIds.Runtime, PageIngredientRelationship.RendersThrough, PageIngredientRequirement.Required));
            graph.Edges.Add(Edge(PublishingPageIngredientIds.PageArtifact, PublishingPageIngredientIds.Layout, PageIngredientRelationship.RendersThrough, PageIngredientRequirement.Required));
            graph.Edges.Add(Edge(PublishingPageIngredientIds.PageArtifact, PublishingPageIngredientIds.ContentType, PageIngredientRelationship.TypedBy, PageIngredientRequirement.Required));
            if (PublishingPageIngredientGraphProjector.UsesTransactionDependencies(revision))
            {
                // v4 models transaction prerequisites rather than aggregate object
                // composition. The target page shell is created first; content,
                // security, fields, Web Parts, and lifecycle are subsequent
                // transactions that require that shell. This lets a deferred
                // optional ingredient prune only its own consumer subtree.
                graph.Edges.Add(Edge(PublishingPageIngredientIds.PublishingContent, PublishingPageIngredientIds.PageArtifact, PageIngredientRelationship.Backs, PageIngredientRequirement.Required));
                graph.Edges.Add(Edge(PublishingPageIngredientIds.Security, PublishingPageIngredientIds.PageArtifact, PageIngredientRelationship.GovernedBy, PageIngredientRequirement.Required));
                graph.Edges.Add(Edge(PublishingPageIngredientIds.Lifecycle, PublishingPageIngredientIds.PageArtifact, PageIngredientRelationship.GovernedBy, PageIngredientRequirement.Required));
                graph.Edges.Add(Edge(PublishingPageIngredientIds.Lifecycle, PublishingPageIngredientIds.PublishingContent, PageIngredientRelationship.DependsOn, PageIngredientRequirement.Required));
            }
            else
            {
                graph.Edges.Add(Edge(PublishingPageIngredientIds.PageArtifact, PublishingPageIngredientIds.PublishingContent, PageIngredientRelationship.Backs, PageIngredientRequirement.Required));
                graph.Edges.Add(Edge(PublishingPageIngredientIds.PageArtifact, PublishingPageIngredientIds.Security, PageIngredientRelationship.GovernedBy, PageIngredientRequirement.Optional));
                graph.Edges.Add(Edge(PublishingPageIngredientIds.PageArtifact, PublishingPageIngredientIds.Lifecycle, PageIngredientRelationship.GovernedBy, PageIngredientRequirement.Required));
            }
        }

        private static void AddFields(
            PublishingPageCaptureBundle snapshot,
            CanonicalPageIngredientGraph graph,
            PublishingPageIngredientGraphProjectionRevision revision)
        {
            var requiredByLayout = new HashSet<string>(
                (snapshot.Layout?.Controls ?? Array.Empty<Layouts.PublishingPageLayoutControl>())
                    .Select(value => value?.FieldName)
                    .Where(value => !string.IsNullOrWhiteSpace(value)),
                StringComparer.OrdinalIgnoreCase);
            foreach (var field in (snapshot.Fields ?? Array.Empty<PageFieldValueSnapshot>())
                         .Where(value => value != null)
                         .OrderBy(value => value.InternalName, StringComparer.OrdinalIgnoreCase))
            {
                var id = PublishingPageIngredientIds.Field(field.InternalName);
                graph.Nodes.Add(Node(
                    id,
                    PageIngredientKind.Field,
                    field.InternalName,
                    field.HasValue || field.Required || requiredByLayout.Contains(field.InternalName),
                    PageIngredientOwnership.Shared,
                    revision == PublishingPageIngredientGraphProjectionRevision.LegacyV1
                        ? "Pages library field schema and ListItem value"
                        : "Pages library ListItem field value; field schema is modeled by the Page Content Type closure",
                    null,
                    snapshot.Runtime?.AdapterId));
                graph.Edges.Add(PublishingPageIngredientGraphProjector.UsesTransactionDependencies(revision)
                    ? Edge(
                        id,
                        PublishingPageIngredientIds.PageArtifact,
                        PageIngredientRelationship.Backs,
                        PageIngredientRequirement.Required)
                    : Edge(
                        PublishingPageIngredientIds.PageArtifact,
                        id,
                        PageIngredientRelationship.Backs,
                        field.Required ? PageIngredientRequirement.Required : PageIngredientRequirement.Optional));
                if (requiredByLayout.Contains(field.InternalName))
                {
                    graph.Edges.Add(Edge(
                        PublishingPageIngredientIds.Layout,
                        id,
                        PageIngredientRelationship.BindsTo,
                        revision == PublishingPageIngredientGraphProjectionRevision.LegacyV1
                            ? PageIngredientRequirement.Required
                            : PageIngredientRequirement.Optional));
                }

                AddTaxonomyRelationships(field, id, snapshot, graph);
            }
        }

        private static void AddTaxonomyRelationships(
            PageFieldValueSnapshot field,
            string fieldIngredientId,
            PublishingPageCaptureBundle snapshot,
            CanonicalPageIngredientGraph graph)
        {
            if (field.Kind != PageFieldValueKind.Taxonomy
                && field.Kind != PageFieldValueKind.TaxonomyCollection)
            {
                return;
            }

            foreach (var value in (field.TaxonomyValues ?? Array.Empty<PageTaxonomyValueSnapshot>())
                         .Where(item => item != null)
                         .OrderBy(item => item.TermGuid, StringComparer.OrdinalIgnoreCase)
                         .ThenBy(item => item.WssId))
            {
                Guid termId;
                Guid.TryParse(value.TermGuid, out termId);
                var relationshipId = PublishingPageIngredientIds.TaxonomyRelationship(field.Id, termId, value.WssId);
                graph.Nodes.Add(Node(
                    relationshipId,
                    PageIngredientKind.Taxonomy,
                    field.InternalName + ":" + value.TermGuid,
                    true,
                    PageIngredientOwnership.Shared,
                    "Taxonomy field binding, live Term resolution, TaxonomyHiddenList and TaxCatchAll relationship",
                    value.Relationship?.EvidenceSha256,
                    snapshot.Runtime?.AdapterId));
                graph.Edges.Add(Edge(
                    fieldIngredientId,
                    relationshipId,
                    PageIngredientRelationship.BindsTo,
                    PageIngredientRequirement.Required));
            }
        }
    }
}
