using PnP.Framework.Migration.Evidence;
using PnP.Framework.Migration.Schema.Fields;
using PnP.Framework.Migration.Taxonomy;
using System;
using System.Collections.Generic;
using System.Linq;

namespace PnP.Framework.Migration.Schema.ContentTypes
{
    public static class ContentTypeSchemaPlanner
    {
        public static ContentTypeMaterializationPlan CreateRequiredClosure(
            ContentTypeSchemaSnapshot schema,
            IEnumerable<TaxonomyTargetMapping> taxonomyMappings = null)
        {
            if (schema == null)
            {
                throw new ArgumentNullException(nameof(schema));
            }

            if (schema.EvidenceState != ContentTypeSchemaEvidenceState.Readable
                || (schema.Availability != EvidenceAvailability.Captured
                    && schema.Availability != EvidenceAvailability.Conflict))
            {
                throw new ArgumentException("Readable content type schema evidence is required.", nameof(schema));
            }

            var mappings = (taxonomyMappings ?? Enumerable.Empty<TaxonomyTargetMapping>()).ToArray();
            var closure = schema.RequiredFieldClosure.ToArray();
            var fields = closure
                .Select(field => CreateFieldPlan(field, closure, mappings))
                .OrderBy(field => field.Role == FieldSchemaRole.Dependency ? 0 : 1)
                .ThenBy(field => field.FieldId)
                .ToList();
            var blocked = fields.Any(field => field.Disposition == FieldSchemaMaterializationDisposition.Block);
            return new ContentTypeMaterializationPlan
            {
                Disposition = blocked ? ContentTypeMaterializationDisposition.Block : ContentTypeMaterializationDisposition.CreateOwned,
                SourceWebUrl = schema.SourceWebUrl,
                ContentTypeId = schema.ContentTypeId,
                Name = schema.Name,
                Description = schema.Description,
                Group = schema.Group,
                ReadOnly = schema.ReadOnly,
                Sealed = schema.Sealed,
                Hidden = schema.Hidden,
                ParentContentTypeId = schema.ParentContentTypeId,
                ParentContentTypeName = schema.ParentContentTypeName,
                RequiredFieldLinks = schema.RequiredFieldLinks.ToList(),
                Fields = fields,
                Reason = blocked
                    ? "The required content type closure contains field schema that needs an explicit capability decision."
                    : "Create or exactly reuse the content type and its required field closure."
            };
        }

        private static FieldSchemaMaterializationPlan CreateFieldPlan(
            FieldSchemaSnapshot field,
            IReadOnlyCollection<FieldSchemaSnapshot> closure,
            IReadOnlyCollection<TaxonomyTargetMapping> mappings)
        {
            var ownership = FieldOwnershipClassifier.Classify(field, closure);
            if (ownership == FieldOwnership.TargetRuntime)
            {
                return Plan(field, ownership, FieldSchemaMaterializationDisposition.RequireTargetRuntime, null, null, null,
                    "The field is supplied by the exact target runtime or parent content type.");
            }

            if (field.ReadOnly || field.Sealed)
            {
                return Plan(field, ownership, FieldSchemaMaterializationDisposition.Block, null, null, null,
                    "A direct required field is read-only or sealed and has no reviewed create-only materialization path.");
            }

            if (field.TypeAsString.StartsWith("Lookup", StringComparison.OrdinalIgnoreCase))
            {
                return Plan(field, ownership, FieldSchemaMaterializationDisposition.Block, null, null, null,
                    "A lookup field requires an explicit target Web, List, and lookup-field mapping.");
            }

            if (field.TypeAsString.StartsWith("TaxonomyFieldType", StringComparison.OrdinalIgnoreCase))
            {
                if (field.Taxonomy == null)
                {
                    return Plan(field, ownership, FieldSchemaMaterializationDisposition.Block, null, null, null,
                        "The taxonomy field schema has no complete source binding.");
                }

                var mapping = mappings.SingleOrDefault(value =>
                    value.SourceTermStoreId == field.Taxonomy.SourceTermStoreId
                    && value.SourceTermSetId == field.Taxonomy.SourceTermSetId);
                if (mapping == null)
                {
                    return Plan(field, ownership, FieldSchemaMaterializationDisposition.Block, null, null, null,
                        $"Taxonomy field requires an explicit target mapping for source store {field.Taxonomy.SourceTermStoreId:D} and term set {field.Taxonomy.SourceTermSetId:D}.");
                }

                var targetSchema = FieldSchemaCanonicalizer.RewriteForTarget(
                    field.SchemaXml,
                    mapping.TargetTermStoreId,
                    mapping.TargetTermSetId,
                    field.Taxonomy.HiddenTextFieldId);
                return Plan(field, ownership, FieldSchemaMaterializationDisposition.CreateOrReuseOwned, targetSchema,
                    mapping.TargetTermStoreId, mapping.TargetTermSetId,
                    "Create or reuse the exact field GUID after rebinding taxonomy to the approved target store and term set.");
            }

            return Plan(field, ownership, FieldSchemaMaterializationDisposition.CreateOrReuseOwned,
                FieldSchemaCanonicalizer.RewriteForTarget(field.SchemaXml), null, null,
                "Create or reuse the exact field GUID from portable source schema.");
        }

        private static FieldSchemaMaterializationPlan Plan(
            FieldSchemaSnapshot field,
            FieldOwnership ownership,
            FieldSchemaMaterializationDisposition disposition,
            string targetSchemaXml,
            Guid? targetTermStoreId,
            Guid? targetTermSetId,
            string reason)
        {
            return new FieldSchemaMaterializationPlan
            {
                FieldId = field.Id,
                InternalName = field.InternalName,
                Title = field.Title,
                TypeAsString = field.TypeAsString,
                Group = field.Group,
                Required = field.Required,
                Hidden = field.Hidden,
                Role = field.Role,
                Ownership = ownership,
                Disposition = disposition,
                SourcePortableSchemaSha256 = field.PortableSchemaSha256,
                TargetSchemaXml = targetSchemaXml,
                TargetPortableSchemaSha256 = targetSchemaXml == null ? null : FieldSchemaCanonicalizer.PortableDigest(targetSchemaXml),
                SourceTermStoreId = field.Taxonomy?.SourceTermStoreId,
                SourceTermSetId = field.Taxonomy?.SourceTermSetId,
                TargetTermStoreId = targetTermStoreId,
                TargetTermSetId = targetTermSetId,
                HiddenTextFieldId = field.Taxonomy?.HiddenTextFieldId,
                Reason = reason
            };
        }
    }
}
