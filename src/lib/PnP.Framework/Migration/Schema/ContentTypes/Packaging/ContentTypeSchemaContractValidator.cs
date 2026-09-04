using PnP.Framework.Migration.Packaging;
using PnP.Framework.Migration.Schema.Fields;
using PnP.Framework.Migration.Taxonomy;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace PnP.Framework.Migration.Schema.ContentTypes.Packaging
{
    internal static class ContentTypeSchemaContractValidator
    {
        public static void ValidateSnapshot(ContentTypeSchemaSnapshot schema)
        {
            if (schema.RequiredFieldLinks == null
                || schema.RequiredFieldClosure == null
                || schema.Sources == null
                || schema.Diagnostics == null)
            {
                throw new InvalidDataException("The content type schema contains a null evidence collection.");
            }

            if (schema.EvidenceState == ContentTypeSchemaEvidenceState.Readable
                && (string.IsNullOrWhiteSpace(schema.ContentTypeId)
                    || string.IsNullOrWhiteSpace(schema.Name)
                    || string.IsNullOrWhiteSpace(schema.ParentContentTypeId)))
            {
                throw new InvalidDataException("Readable content type schema evidence is missing identity or parent information.");
            }

            var duplicateLink = schema.RequiredFieldLinks
                .GroupBy(value => value?.FieldId ?? Guid.Empty)
                .FirstOrDefault(group => group.Key == Guid.Empty || group.Count() > 1);
            var duplicateField = schema.RequiredFieldClosure
                .GroupBy(value => value?.Id ?? Guid.Empty)
                .FirstOrDefault(group => group.Key == Guid.Empty || group.Count() > 1);
            if (duplicateLink != null || duplicateField != null)
            {
                throw new InvalidDataException("The content type schema contains a null, missing-ID, or duplicate field entry.");
            }

            var fieldIds = new HashSet<Guid>(schema.RequiredFieldClosure.Select(value => value.Id));
            if (schema.RequiredFieldLinks.Any(value => !fieldIds.Contains(value.FieldId)))
            {
                throw new InvalidDataException("A required content type field link is absent from the captured schema closure.");
            }

            foreach (var field in schema.RequiredFieldClosure)
            {
                if (field.Sources == null
                    || field.Diagnostics == null
                    || string.IsNullOrWhiteSpace(field.InternalName)
                    || string.IsNullOrWhiteSpace(field.TypeAsString)
                    || string.IsNullOrWhiteSpace(field.SchemaXml)
                    || string.IsNullOrWhiteSpace(field.SchemaXmlSha256)
                    || string.IsNullOrWhiteSpace(field.PortableSchemaSha256))
                {
                    throw new InvalidDataException($"Content type field schema '{field?.Id}' is incomplete.");
                }

                if (!string.Equals(MigrationDigest.ComputeSha256(field.SchemaXml), field.SchemaXmlSha256, StringComparison.OrdinalIgnoreCase)
                    || !string.Equals(FieldSchemaCanonicalizer.PortableDigest(field.SchemaXml), field.PortableSchemaSha256, StringComparison.OrdinalIgnoreCase))
                {
                    throw new InvalidDataException($"Content type field schema digest mismatch: {field.InternalName}");
                }
            }
        }

        public static void ValidatePlan(ContentTypeMaterializationPlan schema)
        {
            if (schema.RequiredFieldLinks == null || schema.Fields == null)
            {
                throw new InvalidDataException("The content type materialization plan contains a null collection.");
            }

            if (string.IsNullOrWhiteSpace(schema.ContentTypeId)
                || string.IsNullOrWhiteSpace(schema.Name)
                || string.IsNullOrWhiteSpace(schema.ParentContentTypeId))
            {
                throw new InvalidDataException("The content type materialization plan is missing identity or parent information.");
            }

            if (schema.Disposition == ContentTypeMaterializationDisposition.ReuseOwned
                && (schema.Fields.Count == 0
                    || schema.Fields.Any(value => value == null
                        || value.Disposition != FieldSchemaMaterializationDisposition.RequireTargetRuntime)))
            {
                throw new InvalidDataException("A target-runtime-only content type plan must contain only explicit target-runtime field requirements.");
            }

            var duplicateLink = schema.RequiredFieldLinks
                .GroupBy(value => value?.FieldId ?? Guid.Empty)
                .FirstOrDefault(group => group.Key == Guid.Empty || group.Count() > 1);
            var duplicateField = schema.Fields
                .GroupBy(value => value?.FieldId ?? Guid.Empty)
                .FirstOrDefault(group => group.Key == Guid.Empty || group.Count() > 1);
            if (duplicateLink != null || duplicateField != null)
            {
                throw new InvalidDataException("The content type materialization plan contains a null, missing-ID, or duplicate field entry.");
            }

            var fieldIds = new HashSet<Guid>(schema.Fields.Select(value => value.FieldId));
            if (schema.RequiredFieldLinks.Any(value => !fieldIds.Contains(value.FieldId)))
            {
                throw new InvalidDataException("A required content type field link is absent from the materialization field closure.");
            }
            if (schema.Fields.Any(value => value.TaxonomyMappingMode == TaxonomyTargetMappingMode.PreserveUnresolvedSourceReference
                && (!value.TypeAsString.StartsWith("TaxonomyFieldType", StringComparison.OrdinalIgnoreCase)
                    || !value.SourceTermSetId.HasValue
                    || !value.TargetTermSetId.HasValue
                    || !value.UnresolvedReferenceTargetVerifiedAbsent
                    || !IsSha256(value.UnresolvedReferenceEvidenceSha256))))
            {
                throw new InvalidDataException("An unresolved taxonomy field plan requires complete source/target identities and digest-sealed target-absence evidence.");
            }
        }

        private static bool IsSha256(string value)
        {
            return !string.IsNullOrWhiteSpace(value)
                && value.Length == 64
                && value.All(character => character >= '0' && character <= '9'
                    || character >= 'a' && character <= 'f'
                    || character >= 'A' && character <= 'F');
        }
    }
}
