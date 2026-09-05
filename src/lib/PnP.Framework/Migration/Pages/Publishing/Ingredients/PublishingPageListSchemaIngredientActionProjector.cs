using PnP.Framework.Migration.Lists.Capture;
using PnP.Framework.Migration.Lists.Planning;
using PnP.Framework.Migration.Pages.Ingredients;
using PnP.Framework.Migration.Pages.Publishing.Capture;
using PnP.Framework.Migration.Pages.Publishing.Planning;
using PnP.Framework.Migration.Schema.ContentTypes;
using PnP.Framework.Migration.Schema.Fields;
using PnP.Framework.Migration.Taxonomy;
using System;
using System.Collections.Generic;
using System.Linq;

namespace PnP.Framework.Migration.Pages.Publishing.Ingredients
{
    internal static class PublishingPageListSchemaIngredientActionProjector
    {
        public static void ProjectSharedClosures(
            PublishingPageCaptureBundle snapshot,
            PublishingPageMigrationPlan plan,
            IDictionary<string, PageIngredientAction> actions)
        {
            var sourceSchemas = snapshot.ListDependencies
                .Where(value => value != null)
                .SelectMany(value => value.SiteContentTypes ?? Array.Empty<ContentTypeSchemaSnapshot>())
                .Where(value => value != null)
                .GroupBy(
                    value => PublishingPageIngredientIds.SiteContentType(SchemaScope(value), value.ContentTypeId),
                    StringComparer.Ordinal)
                .Select(group => group.First())
                .ToArray();
            var plannedSchemas = (plan.ListMigration?.Lists ?? Array.Empty<ListMaterializationPlan>())
                .SelectMany(value => value.SiteContentTypes ?? Array.Empty<ContentTypeClosureNodePlan>())
                .Where(value => value?.Schema != null)
                .ToArray();
            foreach (var sourceSchema in sourceSchemas)
            {
                ProjectSharedContentType(sourceSchema, FindSchemaPlan(sourceSchema, plannedSchemas), actions);
            }
        }

        public static void ProjectList(
            ListDependencySnapshot source,
            ListMaterializationPlan listPlan,
            bool listBlocked,
            IDictionary<string, PageIngredientAction> actions,
            bool transactionDependencyProjection)
        {
            AddListFields(source, listPlan, listBlocked, actions, transactionDependencyProjection);
            AddListContentTypes(source, listPlan, listBlocked, actions, transactionDependencyProjection);
        }

        private static void ProjectSharedContentType(
            ContentTypeSchemaSnapshot sourceSchema,
            ContentTypeClosureNodePlan schemaPlan,
            IDictionary<string, PageIngredientAction> actions)
        {
            var scope = SchemaScope(sourceSchema);
            var blocked = schemaPlan == null || IsContentTypeObjectUnavailable(schemaPlan.Schema);
            PublishingPageIngredientActionFactory.Add(actions, PublishingPageIngredientActionFactory.Create(
                PublishingPageIngredientIds.SiteContentType(scope, sourceSchema.ContentTypeId),
                blocked ? IngredientCapability.Incompatible : IngredientCapability.Available,
                blocked ? IngredientDisposition.Block : IngredientDisposition.Preserve,
                blocked
                    ? "none"
                    : schemaPlan.Schema.Disposition == ContentTypeMaterializationDisposition.ReuseOwned
                        ? "reuse-owned"
                        : "create-owned",
                "policy.site-content-type.closure",
                schemaPlan?.Schema?.Reason ?? "No site Content Type closure materialization decision was produced.",
                blocked ? null : schemaPlan.TargetOwnerWebUrl + "#content-type:" + sourceSchema.ContentTypeId,
                blocked
                    ? null
                    : $"Fresh readback verifies the site Content Type '{sourceSchema.ContentTypeId}', metadata, field links, and ownership."));

            var plannedFields = (schemaPlan?.Schema?.Fields ?? Array.Empty<FieldSchemaMaterializationPlan>())
                .GroupBy(value => value.FieldId)
                .ToDictionary(group => group.Key, group => group.First());
            foreach (var sourceField in sourceSchema.RequiredFieldClosure.Where(value => value != null))
            {
                plannedFields.TryGetValue(sourceField.Id, out var fieldPlan);
                var mapping = fieldPlan == null
                    ? (IngredientCapability.Incompatible, IngredientDisposition.Block, "none")
                    : Map(fieldPlan);
                PublishingPageIngredientActionFactory.Add(actions, PublishingPageIngredientActionFactory.Create(
                    PublishingPageIngredientIds.SiteField(scope, sourceField.Id),
                    mapping.Item1,
                    mapping.Item2,
                    mapping.Item3,
                    "policy.site-field." + (fieldPlan?.Disposition.ToString().ToLowerInvariant() ?? "missing"),
                    fieldPlan?.Reason ?? "No site field-schema materialization decision was produced.",
                    mapping.Item2 == IngredientDisposition.Block
                        ? null
                        : schemaPlan.TargetOwnerWebUrl + "#field:" + sourceField.Id.ToString("D"),
                    mapping.Item2 == IngredientDisposition.Block
                        ? null
                        : $"Fresh readback verifies the portable schema for site field '{sourceField.InternalName}'."));
            }
        }

        private static void AddListFields(
            ListDependencySnapshot source,
            ListMaterializationPlan listPlan,
            bool listBlocked,
            IDictionary<string, PageIngredientAction> actions,
            bool transactionDependencyProjection)
        {
            var plans = listPlan.Fields.ToDictionary(value => value.SourceFieldId);
            foreach (var field in source.Fields.Where(value => value != null))
            {
                plans.TryGetValue(field.Id, out var fieldPlan);
                var mapping = (!transactionDependencyProjection && listBlocked) || fieldPlan == null
                    ? (IngredientCapability.Incompatible, IngredientDisposition.Block, "none")
                    : Map(fieldPlan.Disposition);
                var targetIdentity = listPlan.TargetRootFolderServerRelativeUrl + "#field:" + field.InternalName;
                PublishingPageIngredientActionFactory.Add(actions, PublishingPageIngredientActionFactory.Create(
                    PublishingPageIngredientIds.ListField(source.SourceWebId, source.SourceListId, field.Id),
                    mapping.Item1,
                    mapping.Item2,
                    mapping.Item3,
                    "policy.list-field." + (fieldPlan?.Disposition.ToString().ToLowerInvariant() ?? "missing"),
                    fieldPlan?.Reason ?? "No List field materialization decision was produced.",
                    mapping.Item2 == IngredientDisposition.Drop ? null : targetIdentity,
                    mapping.Item2 == IngredientDisposition.Drop
                        ? "The omitted field schema and all captured raw evidence remain in the immutable snapshot."
                        : mapping.Item2 == IngredientDisposition.Block
                            ? null
                            : $"Fresh List readback verifies the approved schema policy for field '{field.InternalName}'."));
            }
        }

        private static void AddListContentTypes(
            ListDependencySnapshot source,
            ListMaterializationPlan listPlan,
            bool listBlocked,
            IDictionary<string, PageIngredientAction> actions,
            bool transactionDependencyProjection)
        {
            var capturedParents = new HashSet<string>(
                source.SiteContentTypes.Where(value => value != null).Select(value => value.ContentTypeId),
                StringComparer.OrdinalIgnoreCase);
            var droppedFieldIds = new HashSet<Guid>(listPlan.Fields
                .Where(value => value.Disposition == ListFieldMaterializationDisposition.EvidenceOnly)
                .Select(value => value.SourceFieldId));
            foreach (var contentType in source.ContentTypes.Where(value => value != null))
            {
                var missingParent = !string.IsNullOrWhiteSpace(contentType.ParentId)
                    && !ContentTypeRuntimeCatalog.IsTargetRuntime(contentType.ParentId)
                    && !capturedParents.Contains(contentType.ParentId);
                var blocked = (!transactionDependencyProjection && listBlocked) || missingParent;
                var releasedFields = contentType.FieldLinks
                    .Where(value => droppedFieldIds.Contains(value.FieldId))
                    .Select(value => PublishingPageIngredientIds.ListField(source.SourceWebId, source.SourceListId, value.FieldId))
                    .Distinct(StringComparer.Ordinal)
                    .OrderBy(value => value, StringComparer.Ordinal)
                    .ToArray();
                var action = PublishingPageIngredientActionFactory.Create(
                    PublishingPageIngredientIds.ListContentType(source.SourceWebId, source.SourceListId, contentType.Id),
                    blocked ? IngredientCapability.Incompatible : IngredientCapability.Available,
                    blocked
                        ? IngredientDisposition.Block
                        : releasedFields.Length > 0 ? IngredientDisposition.Transform : IngredientDisposition.Preserve,
                    blocked
                        ? "none"
                        : releasedFields.Length > 0
                            ? "materialize-list-content-type-with-runtime-cache-links-released"
                            : "materialize-list-content-type-membership",
                    "policy.list-content-type.membership",
                    missingParent
                        ? "The List-local Content Type references a custom parent whose exact site Content Type closure is absent."
                        : listBlocked && !transactionDependencyProjection
                            ? "The owning List has no executable materialization plan."
                        : releasedFields.Length > 0
                            ? "Create or reuse the captured List content type while explicitly releasing SharePoint-owned taxonomy cache FieldLinks; their source schema and values remain in the snapshot."
                            : "Create or reuse the captured List content type membership and apply its field links and ordering.",
                    blocked ? null : listPlan.TargetRootFolderServerRelativeUrl + "#content-type:" + contentType.Id,
                    blocked
                        ? null
                        : $"The List receipt maps source Content Type '{contentType.Id}' to a verified target Content Type ID.");
                foreach (var releasedField in releasedFields)
                {
                    action.ReleasedDependencyIngredientIds.Add(releasedField);
                }
                PublishingPageIngredientActionFactory.Add(actions, action);
            }
        }

        private static bool IsContentTypeObjectUnavailable(ContentTypeMaterializationPlan plan)
        {
            if (plan == null)
            {
                return true;
            }

            return plan.Disposition == ContentTypeMaterializationDisposition.Block
                && !(plan.Fields?.Any(value => value?.Disposition == FieldSchemaMaterializationDisposition.Block) ?? false);
        }

        private static (IngredientCapability Capability, IngredientDisposition Disposition, string Realization) Map(
            ListFieldMaterializationDisposition disposition)
        {
            switch (disposition)
            {
                case ListFieldMaterializationDisposition.RequireTargetRuntime:
                case ListFieldMaterializationDisposition.RequireTargetRuntimeAndCopyValue:
                    return (IngredientCapability.Available, IngredientDisposition.Substitute, "reuse-target-runtime-schema");
                case ListFieldMaterializationDisposition.CreateOrReuseOwnedAndCopyValue:
                    return (IngredientCapability.Available, IngredientDisposition.Preserve, "create-or-reuse-owned-and-copy-values");
                case ListFieldMaterializationDisposition.CreateOrReuseOwnedCalculated:
                    return (IngredientCapability.Available, IngredientDisposition.Preserve, "create-or-reuse-owned-calculated-schema");
                case ListFieldMaterializationDisposition.MapLookup:
                    return (IngredientCapability.Available, IngredientDisposition.Transform, "map-lookup-list-and-item-identities");
                case ListFieldMaterializationDisposition.MapTaxonomy:
                    return (IngredientCapability.Available, IngredientDisposition.Transform, "map-taxonomy-store-and-set");
                case ListFieldMaterializationDisposition.CreateOrReuseOwnedSchemaOnly:
                    return (IngredientCapability.Available, IngredientDisposition.Preserve, "create-or-reuse-owned-schema-only");
                case ListFieldMaterializationDisposition.EvidenceOnly:
                    return (IngredientCapability.Unknown, IngredientDisposition.Drop, "retain-snapshot-only");
                default:
                    return (IngredientCapability.Incompatible, IngredientDisposition.Block, "none");
            }
        }

        private static (IngredientCapability Capability, IngredientDisposition Disposition, string Realization) Map(
            FieldSchemaMaterializationDisposition disposition)
        {
            switch (disposition)
            {
                case FieldSchemaMaterializationDisposition.RequireTargetRuntime:
                    return (IngredientCapability.Available, IngredientDisposition.Substitute, "reuse-target-runtime-schema");
                case FieldSchemaMaterializationDisposition.CreateOrReuseOwned:
                    return (IngredientCapability.Available, IngredientDisposition.Preserve, "create-or-reuse-owned-schema");
                default:
                    return (IngredientCapability.Incompatible, IngredientDisposition.Block, "none");
            }
        }

        private static (IngredientCapability Capability, IngredientDisposition Disposition, string Realization) Map(
            FieldSchemaMaterializationPlan field)
        {
            return field.TaxonomyMappingMode == TaxonomyTargetMappingMode.PreserveUnresolvedSourceReference
                ? (IngredientCapability.Available, IngredientDisposition.Transform, "create-schema-preserving-unresolved-taxonomy-reference")
                : Map(field.Disposition);
        }

        private static ContentTypeClosureNodePlan FindSchemaPlan(
            ContentTypeSchemaSnapshot source,
            IEnumerable<ContentTypeClosureNodePlan> plans)
        {
            var candidates = plans.Where(value => string.Equals(
                    value.Schema.ContentTypeId,
                    source.ContentTypeId,
                    StringComparison.OrdinalIgnoreCase))
                .ToArray();
            var sourceScope = SchemaScope(source);
            var exact = candidates.FirstOrDefault(value => string.Equals(
                UrlScope(value.SourceOwnerWebUrl),
                sourceScope,
                StringComparison.OrdinalIgnoreCase));
            return exact ?? (candidates.Length == 1 ? candidates[0] : null);
        }

        private static string SchemaScope(ContentTypeSchemaSnapshot schema)
        {
            if (!string.IsNullOrWhiteSpace(schema?.SourceScope))
            {
                if (Uri.TryCreate(schema.SourceScope, UriKind.Absolute, out var absoluteScope))
                {
                    return NormalizeScope(absoluteScope.AbsolutePath);
                }
                return NormalizeScope(schema.SourceScope);
            }
            return UrlScope(schema?.SourceWebUrl);
        }

        private static string UrlScope(string value)
        {
            return Uri.TryCreate(value, UriKind.Absolute, out var absolute)
                ? NormalizeScope(absolute.AbsolutePath)
                : NormalizeScope(value);
        }

        private static string NormalizeScope(string value)
        {
            var normalized = Uri.UnescapeDataString(value ?? string.Empty).Replace('\\', '/').TrimEnd('/');
            return normalized.Length == 0 ? "/" : normalized;
        }
    }
}
