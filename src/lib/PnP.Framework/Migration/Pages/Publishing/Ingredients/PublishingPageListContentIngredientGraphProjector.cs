using PnP.Framework.Migration.Lists.Capture;
using PnP.Framework.Migration.Lists.Items;
using PnP.Framework.Migration.Lists.Views;
using PnP.Framework.Migration.Pages.Ingredients;
using System;
using System.Collections.Generic;
using System.Linq;
using static PnP.Framework.Migration.Pages.Publishing.Ingredients.PublishingPageIngredientGraphFactory;

namespace PnP.Framework.Migration.Pages.Publishing.Ingredients
{
    internal static class PublishingPageListContentIngredientGraphProjector
    {
        public static void Project(
            ListDependencySnapshot list,
            string listId,
            CanonicalPageIngredientGraph graph,
            PublishingPageIngredientGraphProjectionRevision revision,
            string resourceOwnerWebId)
        {
            var fieldIdsByName = list.Fields.Where(value => value != null).ToDictionary(
                value => value.InternalName,
                value => PublishingPageIngredientIds.ListField(list.SourceWebId, list.SourceListId, value.Id),
                StringComparer.OrdinalIgnoreCase);
            var contentTypeIds = list.ContentTypes.Where(value => value != null).ToDictionary(
                value => value.Id,
                value => PublishingPageIngredientIds.ListContentType(list.SourceWebId, list.SourceListId, value.Id),
                StringComparer.OrdinalIgnoreCase);
            foreach (var item in list.Items.Where(value => value != null).OrderBy(value => value.SourceItemId))
            {
                AddItem(list, item, listId, fieldIdsByName, contentTypeIds, graph, revision);
            }
            var renderingResources = list.ViewRenderingResources
                .Where(value => value != null && !string.IsNullOrWhiteSpace(value.Id))
                .GroupBy(value => value.Id, StringComparer.Ordinal)
                .ToDictionary(group => group.Key, group => group.First(), StringComparer.Ordinal);
            foreach (var resource in renderingResources.Values.OrderBy(value => value.Id, StringComparer.Ordinal))
            {
                var resourceId = PublishingPageIngredientIds.ViewRenderingResource(list.SourceSiteId, resource.Id);
                graph.Nodes.Add(Node(
                    resourceId,
                    PageIngredientKind.Asset,
                    resource.SourceServerRelativeUrl ?? resource.SourceAbsoluteUrl,
                    true,
                    PageIngredientOwnership.Shared,
                    "Captured external content required by one or more List Views",
                    resource.Artifact?.Sha256,
                    null));
                if (PublishingPageIngredientGraphProjector.UsesOwnerWebDependencies(revision)
                    && !string.IsNullOrWhiteSpace(resourceOwnerWebId))
                {
                    graph.Edges.Add(Edge(
                        resourceId,
                        resourceOwnerWebId,
                        PageIngredientRelationship.DependsOn,
                        PageIngredientRequirement.Required));
                }
                else if (PublishingPageIngredientGraphProjector.UsesTransactionDependencies(revision))
                {
                    graph.Edges.Add(Edge(
                        resourceId,
                        listId,
                        PageIngredientRelationship.DependsOn,
                        PageIngredientRequirement.Required));
                }
            }
            foreach (var view in list.Views.Where(value => value != null).OrderBy(value => value.Id))
            {
                var viewId = PublishingPageIngredientIds.View(list.SourceWebId, list.SourceListId, view.Id);
                graph.Nodes.Add(Node(
                    viewId,
                    PageIngredientKind.View,
                    view.Title,
                    true,
                    PageIngredientOwnership.Shared,
                    "Captured List View",
                    view.ListViewXmlSha256,
                    null));
                graph.Edges.Add(Edge(listId, viewId, PageIngredientRelationship.Backs, PageIngredientRequirement.Optional));
                if (PublishingPageIngredientGraphProjector.UsesTransactionDependencies(revision))
                {
                    graph.Edges.Add(Edge(
                        viewId,
                        listId,
                        PageIngredientRelationship.DependsOn,
                        PageIngredientRequirement.Required));
                }
                foreach (var binding in (view.RenderingResourceBindings ?? Array.Empty<ListViewRenderingResourceBindingSnapshot>())
                             .Where(value => value != null && renderingResources.ContainsKey(value.ResourceId ?? string.Empty))
                             .GroupBy(value => value.ResourceId, StringComparer.Ordinal)
                             .Select(group => group.First()))
                {
                    graph.Edges.Add(Edge(
                        viewId,
                        PublishingPageIngredientIds.ViewRenderingResource(list.SourceSiteId, binding.ResourceId),
                        PageIngredientRelationship.DependsOn,
                        PageIngredientRequirement.Required));
                }
                foreach (var fieldName in view.ViewFields
                             .Where(value => !string.IsNullOrWhiteSpace(value))
                             .Distinct(StringComparer.OrdinalIgnoreCase))
                {
                    if (fieldIdsByName.TryGetValue(fieldName, out var fieldId))
                    {
                        graph.Edges.Add(Edge(viewId, fieldId, PageIngredientRelationship.BindsTo, PageIngredientRequirement.Required));
                    }
                }
            }
        }

        private static void AddItem(
            ListDependencySnapshot list,
            ListItemSnapshot item,
            string listId,
            IDictionary<string, string> fieldIdsByName,
            IDictionary<string, string> contentTypeIds,
            CanonicalPageIngredientGraph graph,
            PublishingPageIngredientGraphProjectionRevision revision)
        {
            var itemId = PublishingPageIngredientIds.ListItem(list.SourceWebId, list.SourceListId, item.SourceItemId);
            graph.Nodes.Add(Node(
                itemId,
                PageIngredientKind.ListItem,
                list.Title + " item " + item.SourceItemId,
                true,
                PageIngredientOwnership.SourceOwned,
                "Captured current List item state",
                null,
                null));
            graph.Edges.Add(Edge(listId, itemId, PageIngredientRelationship.Backs, PageIngredientRequirement.Optional));
            if (PublishingPageIngredientGraphProjector.UsesTransactionDependencies(revision))
            {
                graph.Edges.Add(Edge(
                    itemId,
                    listId,
                    PageIngredientRelationship.DependsOn,
                    PageIngredientRequirement.Required));
            }
            foreach (var value in item.Values.Where(value => value != null && value.Kind != ListItemValueKind.Null))
            {
                // v1/v2 encoded one item-to-field edge for every non-null value. Large
                // captured Lists therefore produced an O(items * fields) graph even though
                // the List transaction already owns the union of value-bearing fields.
                // v3 keeps that exact schema dependency once at List scope; the item action
                // still evaluates every captured value independently.
                if ((revision == PublishingPageIngredientGraphProjectionRevision.LegacyV1
                        || revision == PublishingPageIngredientGraphProjectionRevision.Version2)
                    && fieldIdsByName.TryGetValue(value.InternalName, out var fieldId))
                {
                    graph.Edges.Add(Edge(itemId, fieldId, PageIngredientRelationship.BindsTo, PageIngredientRequirement.Required));
                }
                if (string.Equals(value.InternalName, "ContentTypeId", StringComparison.OrdinalIgnoreCase))
                {
                    var sourceContentTypeId = value.ScalarValue ?? value.RawValue;
                    if (!string.IsNullOrWhiteSpace(sourceContentTypeId)
                        && contentTypeIds.TryGetValue(sourceContentTypeId, out var contentTypeId))
                    {
                        graph.Edges.Add(Edge(itemId, contentTypeId, PageIngredientRelationship.TypedBy, PageIngredientRequirement.Required));
                    }
                }
            }

            if (item.Document != null)
            {
                var documentId = PublishingPageIngredientIds.ListDocument(list.SourceWebId, list.SourceListId, item.SourceItemId);
                graph.Nodes.Add(Node(
                    documentId,
                    PageIngredientKind.Document,
                    item.Document.ServerRelativeUrl,
                    true,
                    PageIngredientOwnership.SourceOwned,
                    "Captured current List document or folder",
                    item.Document.Content?.Artifact?.Sha256,
                    null));
                graph.Edges.Add(Edge(itemId, documentId, PageIngredientRelationship.Backs, PageIngredientRequirement.Required));
                if (PublishingPageIngredientGraphProjector.UsesTransactionDependencies(revision))
                {
                    graph.Edges.Add(Edge(
                        documentId,
                        listId,
                        PageIngredientRelationship.DependsOn,
                        PageIngredientRequirement.Required));
                }
                if (item.Document.InformationProtection != null)
                {
                    var informationProtectionId = PublishingPageIngredientIds.ListDocumentInformationProtection(
                        list.SourceWebId,
                        list.SourceListId,
                        item.SourceItemId);
                    graph.Nodes.Add(Node(
                        informationProtectionId,
                        PageIngredientKind.Policy,
                        item.Document.InformationProtection.LabelId,
                        true,
                        PageIngredientOwnership.Shared,
                        "Captured document-level Microsoft Information Protection assignment",
                        item.Document.InformationProtection.LabelHash,
                        null));
                    graph.Edges.Add(Edge(
                        documentId,
                        informationProtectionId,
                        PageIngredientRelationship.GovernedBy,
                        revision == PublishingPageIngredientGraphProjectionRevision.Version6
                            || revision == PublishingPageIngredientGraphProjectionRevision.CurrentV7
                            ? PageIngredientRequirement.Optional
                            : PageIngredientRequirement.Required));
                    if (PublishingPageIngredientGraphProjector.UsesTransactionDependencies(revision))
                    {
                        graph.Edges.Add(Edge(
                            informationProtectionId,
                            documentId,
                            PageIngredientRelationship.DependsOn,
                            PageIngredientRequirement.Required));
                    }
                }
            }

            foreach (var attachment in item.Attachments.Where(value => value != null).OrderBy(value => value.FileName, StringComparer.OrdinalIgnoreCase))
            {
                var attachmentId = PublishingPageIngredientIds.ListAttachment(
                    list.SourceWebId,
                    list.SourceListId,
                    item.SourceItemId,
                    attachment.FileName);
                graph.Nodes.Add(Node(
                    attachmentId,
                    PageIngredientKind.Attachment,
                    attachment.ServerRelativeUrl,
                    true,
                    PageIngredientOwnership.SourceOwned,
                    "Captured List item attachment",
                    attachment.Content?.Artifact?.Sha256,
                    null));
                graph.Edges.Add(PublishingPageIngredientGraphProjector.UsesTransactionDependencies(revision)
                    ? Edge(
                        attachmentId,
                        itemId,
                        PageIngredientRelationship.DependsOn,
                        PageIngredientRequirement.Required)
                    : Edge(
                        itemId,
                        attachmentId,
                        PageIngredientRelationship.Backs,
                        PageIngredientRequirement.Required));
            }
        }
    }
}
