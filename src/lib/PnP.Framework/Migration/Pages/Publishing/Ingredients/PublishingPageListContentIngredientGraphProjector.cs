using PnP.Framework.Migration.Lists.Capture;
using PnP.Framework.Migration.Lists.Items;
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
            CanonicalPageIngredientGraph graph)
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
                AddItem(list, item, listId, fieldIdsByName, contentTypeIds, graph);
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
            CanonicalPageIngredientGraph graph)
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
            foreach (var value in item.Values.Where(value => value != null && value.Kind != ListItemValueKind.Null))
            {
                if (fieldIdsByName.TryGetValue(value.InternalName, out var fieldId))
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
                graph.Edges.Add(Edge(itemId, attachmentId, PageIngredientRelationship.Backs, PageIngredientRequirement.Required));
            }
        }
    }
}
