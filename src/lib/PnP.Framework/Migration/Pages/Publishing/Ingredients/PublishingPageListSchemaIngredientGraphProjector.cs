using PnP.Framework.Migration.Lists.Capture;
using PnP.Framework.Migration.Lists.Items;
using PnP.Framework.Migration.Pages.Ingredients;
using PnP.Framework.Migration.Pages.Publishing.Capture;
using PnP.Framework.Migration.Schema.ContentTypes;
using System;
using System.Collections.Generic;
using System.Linq;
using static PnP.Framework.Migration.Pages.Publishing.Ingredients.PublishingPageIngredientGraphFactory;

namespace PnP.Framework.Migration.Pages.Publishing.Ingredients
{
    internal static class PublishingPageListSchemaIngredientGraphProjector
    {
        public static void ProjectSharedClosures(
            PublishingPageCaptureBundle snapshot,
            CanonicalPageIngredientGraph graph,
            PublishingPageIngredientGraphProjectionRevision revision)
        {
            var schemas = snapshot.ListDependencies
                .Where(value => value != null)
                .SelectMany(value => value.SiteContentTypes ?? Array.Empty<ContentTypeSchemaSnapshot>())
                .Where(value => value != null)
                .GroupBy(
                    value => PublishingPageIngredientIds.SiteContentType(SchemaScope(value), value.ContentTypeId),
                    StringComparer.Ordinal)
                .Select(group => group.First())
                .OrderBy(value => SchemaScope(value), StringComparer.OrdinalIgnoreCase)
                .ThenBy(value => value.ContentTypeId, StringComparer.OrdinalIgnoreCase)
                .ToArray();
            var schemaIds = new HashSet<string>(schemas.Select(value =>
                PublishingPageIngredientIds.SiteContentType(SchemaScope(value), value.ContentTypeId)), StringComparer.Ordinal);
            var addedFields = new HashSet<string>(StringComparer.Ordinal);
            foreach (var schema in schemas)
            {
                AddSharedContentType(snapshot, schema, schemaIds, addedFields, graph, revision);
            }
        }

        public static void ProjectList(
            ListDependencySnapshot list,
            string listId,
            IDictionary<Guid, ListDependencySnapshot> listsById,
            CanonicalPageIngredientGraph graph,
            PublishingPageIngredientGraphProjectionRevision revision)
        {
            var fieldsUsedByCapturedItems = revision == PublishingPageIngredientGraphProjectionRevision.Version3
                || PublishingPageIngredientGraphProjector.UsesTransactionDependencies(revision)
                ? new HashSet<string>(
                    (list.Items ?? Array.Empty<ListItemSnapshot>())
                        .Where(item => item != null)
                        .SelectMany(item => item.Values ?? Array.Empty<ListItemValueSnapshot>())
                        .Where(value => value != null
                            && value.Kind != ListItemValueKind.Null
                            && !string.IsNullOrWhiteSpace(value.InternalName))
                        .Select(value => value.InternalName),
                    StringComparer.OrdinalIgnoreCase)
                : null;
            var fieldIdsByName = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            foreach (var field in list.Fields.Where(value => value != null).OrderBy(value => value.Id))
            {
                var fieldId = PublishingPageIngredientIds.ListField(list.SourceWebId, list.SourceListId, field.Id);
                fieldIdsByName[field.InternalName] = fieldId;
                graph.Nodes.Add(Node(
                    fieldId,
                    PageIngredientKind.Field,
                    field.InternalName,
                    true,
                    PageIngredientOwnership.Shared,
                    "Captured List field schema",
                    field.PortableSchemaSha256,
                    null));
                graph.Edges.Add(Edge(
                    listId,
                    fieldId,
                    PageIngredientRelationship.Backs,
                    PublishingPageIngredientGraphProjector.UsesTransactionDependencies(revision)
                        ? PageIngredientRequirement.Conditional
                        : fieldsUsedByCapturedItems?.Contains(field.InternalName) == true
                            ? PageIngredientRequirement.Required
                            : PageIngredientRequirement.Conditional));
                if (PublishingPageIngredientGraphProjector.UsesTransactionDependencies(revision))
                {
                    graph.Edges.Add(Edge(
                        fieldId,
                        listId,
                        PageIngredientRelationship.DependsOn,
                        PageIngredientRequirement.Required));
                }
                if (field.SourceLookupListId.HasValue)
                {
                    if (listsById.TryGetValue(field.SourceLookupListId.Value, out var provider))
                    {
                        graph.Edges.Add(Edge(
                            fieldId,
                            PublishingPageIngredientIds.List(provider.SourceWebId, provider.SourceListId),
                            PageIngredientRelationship.DependsOn,
                            PageIngredientRequirement.Required));
                    }
                }
            }

            foreach (var contentType in list.ContentTypes.Where(value => value != null).OrderBy(value => value.Id, StringComparer.OrdinalIgnoreCase))
            {
                var contentTypeId = PublishingPageIngredientIds.ListContentType(list.SourceWebId, list.SourceListId, contentType.Id);
                graph.Nodes.Add(Node(
                    contentTypeId,
                    PageIngredientKind.ContentType,
                    contentType.Name,
                    true,
                    PageIngredientOwnership.Shared,
                    "Captured List-local Content Type",
                    null,
                    null));
                graph.Edges.Add(Edge(listId, contentTypeId, PageIngredientRelationship.TypedBy, PageIngredientRequirement.Conditional));
                if (PublishingPageIngredientGraphProjector.UsesTransactionDependencies(revision))
                {
                    graph.Edges.Add(Edge(
                        contentTypeId,
                        listId,
                        PageIngredientRelationship.DependsOn,
                        PageIngredientRequirement.Required));
                }
                var parent = list.SiteContentTypes.FirstOrDefault(value => value != null
                    && string.Equals(value.ContentTypeId, contentType.ParentId, StringComparison.OrdinalIgnoreCase));
                if (parent != null)
                {
                    graph.Edges.Add(Edge(
                        contentTypeId,
                        PublishingPageIngredientIds.SiteContentType(SchemaScope(parent), parent.ContentTypeId),
                        PageIngredientRelationship.TypedBy,
                        PageIngredientRequirement.Required));
                }
                foreach (var link in contentType.FieldLinks
                             .Where(value => value != null && fieldIdsByName.ContainsKey(value.InternalName))
                             .GroupBy(value => value.InternalName, StringComparer.OrdinalIgnoreCase)
                             .Select(group => group.First()))
                {
                    graph.Edges.Add(Edge(
                        contentTypeId,
                        fieldIdsByName[link.InternalName],
                        PageIngredientRelationship.BindsTo,
                        PageIngredientRequirement.Required));
                }
            }
        }

        public static string SchemaScope(ContentTypeSchemaSnapshot schema)
        {
            if (!string.IsNullOrWhiteSpace(schema?.SourceScope))
            {
                if (Uri.TryCreate(schema.SourceScope, UriKind.Absolute, out var absoluteScope))
                {
                    return NormalizeScope(absoluteScope.AbsolutePath);
                }
                return NormalizeScope(schema.SourceScope);
            }
            if (Uri.TryCreate(schema?.SourceWebUrl, UriKind.Absolute, out var sourceWeb))
            {
                return NormalizeScope(sourceWeb.AbsolutePath);
            }
            return "/";
        }

        private static void AddSharedContentType(
            PublishingPageCaptureBundle snapshot,
            ContentTypeSchemaSnapshot schema,
            ISet<string> schemaIds,
            ISet<string> addedFields,
            CanonicalPageIngredientGraph graph,
            PublishingPageIngredientGraphProjectionRevision revision)
        {
            var scope = SchemaScope(schema);
            var contentTypeId = PublishingPageIngredientIds.SiteContentType(scope, schema.ContentTypeId);
            var ownerWebId = PublishingPageIngredientGraphProjector.UsesOwnerWebDependencies(revision)
                ? PublishingPageIngredientOwnerWebResolver.ExactOrContaining(
                    snapshot,
                    schema.SourceScope ?? schema.SourceWebUrl)
                : null;
            graph.Nodes.Add(Node(
                contentTypeId,
                PageIngredientKind.ContentType,
                schema.Name,
                true,
                PageIngredientOwnership.Shared,
                "Captured site Content Type closure",
                null,
                null));
            if (!string.IsNullOrWhiteSpace(ownerWebId))
            {
                graph.Edges.Add(Edge(
                    contentTypeId,
                    ownerWebId,
                    PageIngredientRelationship.DependsOn,
                    PageIngredientRequirement.Required));
            }

            var parentId = PublishingPageIngredientIds.SiteContentType(scope, schema.ParentContentTypeId);
            if (schemaIds.Contains(parentId))
            {
                graph.Edges.Add(Edge(
                    contentTypeId,
                    parentId,
                    PageIngredientRelationship.TypedBy,
                    PageIngredientRequirement.Required));
            }

            foreach (var field in schema.RequiredFieldClosure
                         .Where(value => value != null)
                         .GroupBy(value => value.Id)
                         .Select(group => group.First())
                         .OrderBy(value => value.Id))
            {
                var fieldId = PublishingPageIngredientIds.SiteField(scope, field.Id);
                if (addedFields.Add(fieldId))
                {
                    graph.Nodes.Add(Node(
                        fieldId,
                        PageIngredientKind.Field,
                        field.InternalName,
                        true,
                        PageIngredientOwnership.Shared,
                        "Captured site field-schema closure",
                        field.PortableSchemaSha256,
                        null));
                    if (!string.IsNullOrWhiteSpace(ownerWebId))
                    {
                        graph.Edges.Add(Edge(
                            fieldId,
                            ownerWebId,
                            PageIngredientRelationship.DependsOn,
                            PageIngredientRequirement.Required));
                    }
                }
                graph.Edges.Add(Edge(
                    contentTypeId,
                    fieldId,
                    PageIngredientRelationship.BindsTo,
                    PageIngredientRequirement.Required));
            }
        }

        private static string NormalizeScope(string value)
        {
            var normalized = Uri.UnescapeDataString(value ?? string.Empty).Replace('\\', '/').TrimEnd('/');
            return normalized.Length == 0 ? "/" : normalized;
        }
    }
}
