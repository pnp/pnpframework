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
            return CreatePlan(
                schema,
                fields,
                blocked ? ContentTypeMaterializationDisposition.Block : ContentTypeMaterializationDisposition.CreateOwned,
                blocked
                    ? "The required content type closure contains field schema that needs an explicit capability decision."
                    : "Create or exactly reuse the content type and its required field closure.");
        }

        public static bool TryCreateTargetRuntimeRequirement(
            ContentTypeSchemaSnapshot schema,
            out ContentTypeMaterializationPlan plan)
        {
            plan = null;
            if (schema == null
                || schema.EvidenceState != ContentTypeSchemaEvidenceState.Partial
                || schema.Availability != EvidenceAvailability.Partial
                || string.IsNullOrWhiteSpace(schema.ContentTypeId)
                || string.IsNullOrWhiteSpace(schema.Name)
                || string.IsNullOrWhiteSpace(schema.ParentContentTypeId)
                || schema.RequiredFieldLinks == null
                || schema.RequiredFieldClosure == null
                || schema.RequiredFieldClosure.Count == 0
                || schema.RequiredFieldLinks.Any(value => value == null || value.FieldId == Guid.Empty)
                || schema.RequiredFieldClosure.Any(value => value == null
                    || value.Id == Guid.Empty
                    || string.IsNullOrWhiteSpace(value.InternalName)
                    || string.IsNullOrWhiteSpace(value.TypeAsString)
                    || string.IsNullOrWhiteSpace(value.SchemaXml)))
            {
                return false;
            }

            var closure = schema.RequiredFieldClosure.ToArray();
            var closureIds = new HashSet<Guid>(closure.Select(value => value.Id));
            if (closureIds.Count != closure.Length
                || schema.RequiredFieldLinks.Select(value => value.FieldId).Distinct().Count() != schema.RequiredFieldLinks.Count
                || schema.RequiredFieldLinks.Any(value => !closureIds.Contains(value.FieldId)))
            {
                return false;
            }

            var fields = new List<FieldSchemaMaterializationPlan>();
            foreach (var field in closure)
            {
                var ownership = FieldOwnershipClassifier.Classify(field, closure);
                if (ownership != FieldOwnership.TargetRuntime)
                {
                    return false;
                }

                fields.Add(Plan(
                    field,
                    ownership,
                    FieldSchemaMaterializationDisposition.RequireTargetRuntime,
                    null,
                    null,
                    null,
                    "The partial source closure identifies this as a target-runtime field; require its exact ID, internal name, and type without creating or repairing schema."));
            }

            plan = CreatePlan(
                schema,
                fields
                    .OrderBy(field => field.Role == FieldSchemaRole.Dependency ? 0 : 1)
                    .ThenBy(field => field.FieldId)
                    .ToList(),
                ContentTypeMaterializationDisposition.ReuseOwned,
                "Source content type evidence is partial, so creation is forbidden. Require the exact existing target-runtime content type, parent lineage, metadata, captured field links, and captured target-runtime field closure.");
            return true;
        }

        private static ContentTypeMaterializationPlan CreatePlan(
            ContentTypeSchemaSnapshot schema,
            IList<FieldSchemaMaterializationPlan> fields,
            ContentTypeMaterializationDisposition disposition,
            string reason)
        {
            return new ContentTypeMaterializationPlan
            {
                Disposition = disposition,
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
                Reason = reason
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
                var featureRuntimeFallback = ContentTypeRuntimeCatalog.IsDocumentIdField(field.Id)
                    ? FieldSchemaCanonicalizer.RewriteForTarget(field.SchemaXml)
                    : null;
                return Plan(field, ownership, FieldSchemaMaterializationDisposition.RequireTargetRuntime, featureRuntimeFallback, null, null,
                    featureRuntimeFallback == null
                        ? "The field is supplied by the exact target runtime or parent content type."
                        : "Prefer the Document ID feature runtime field; retain the exact captured schema as a sealed fallback because multi-Web site activation completes field registration asynchronously.");
            }

            if (field.TypeAsString.StartsWith("Calculated", StringComparison.OrdinalIgnoreCase))
            {
                return Plan(field, ownership, FieldSchemaMaterializationDisposition.CreateOrReuseOwned,
                    FieldSchemaCanonicalizer.RewriteForTarget(field.SchemaXml), null, null,
                    "Create or reuse the exact source-owned calculated field schema; target SharePoint computes its value.");
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
                if (mapping.Mode == TaxonomyTargetMappingMode.PreserveUnresolvedSourceReference
                    && (!mapping.UnresolvedReferenceTargetVerifiedAbsent
                        || !IsSha256(mapping.UnresolvedReferenceEvidenceSha256)))
                {
                    return Plan(field, ownership, FieldSchemaMaterializationDisposition.Block, null, null, null,
                        "An unresolved taxonomy reference requires digest-sealed evidence that the selected target TermSet GUID is absent.");
                }

                var targetSchema = FieldSchemaCanonicalizer.RewriteForTarget(
                    field.SchemaXml,
                    mapping.TargetTermStoreId,
                    mapping.TargetTermSetId,
                    field.Taxonomy.HiddenTextFieldId);
                var planned = Plan(field, ownership, FieldSchemaMaterializationDisposition.CreateOrReuseOwned, targetSchema,
                    mapping.TargetTermStoreId, mapping.TargetTermSetId,
                    mapping.Mode == TaxonomyTargetMappingMode.PreserveUnresolvedSourceReference
                        ? mapping.TargetTermSetId == field.Taxonomy.SourceTermSetId
                            ? "Create or reuse the exact field GUID in the target Term Store while preserving the source-invalid TermSet GUID as an unresolved reference. Do not create, substitute, or repair the missing TermSet."
                            : $"Create or reuse the exact field GUID with digest-derived absent target TermSet '{mapping.TargetTermSetId:D}' because source GUID '{field.Taxonomy.SourceTermSetId:D}' is already live at the target. Preserve the original source GUID in the sealed plan and do not create, substitute, or repair a TermSet asset."
                        : "Create or reuse the exact field GUID after rebinding taxonomy to the selected target store and term set. Execution remains gated by reviewed taxonomy-asset admission and fresh readback.");
                planned.TaxonomyMappingMode = mapping.Mode;
                planned.UnresolvedReferenceTargetVerifiedAbsent = mapping.UnresolvedReferenceTargetVerifiedAbsent;
                planned.UnresolvedReferenceEvidenceSha256 = mapping.UnresolvedReferenceEvidenceSha256;
                return planned;
            }

            return Plan(field, ownership, FieldSchemaMaterializationDisposition.CreateOrReuseOwned,
                FieldSchemaCanonicalizer.RewriteForTarget(field.SchemaXml), null, null,
                "Create or reuse the exact field GUID from portable source schema.");
        }

        private static bool IsSha256(string value)
        {
            return !string.IsNullOrWhiteSpace(value)
                && value.Length == 64
                && value.All(character => character >= '0' && character <= '9'
                    || character >= 'a' && character <= 'f'
                    || character >= 'A' && character <= 'F');
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
                SourceAnchorTermId = field.Taxonomy?.AnchorTermId,
                TargetTermStoreId = targetTermStoreId,
                TargetTermSetId = targetTermSetId,
                TargetAnchorTermId = field.Taxonomy?.AnchorTermId,
                HiddenTextFieldId = field.Taxonomy?.HiddenTextFieldId,
                Reason = reason
            };
        }
    }
}
