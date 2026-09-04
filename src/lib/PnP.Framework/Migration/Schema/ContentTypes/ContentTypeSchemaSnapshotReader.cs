using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.Taxonomy;
using PnP.Framework.Migration.Evidence;
using PnP.Framework.Migration.Packaging;
using PnP.Framework.Migration.Schema.Fields;
using System;
using System.Collections.Generic;
using System.Linq;

namespace PnP.Framework.Migration.Schema.ContentTypes
{
    internal static class ContentTypeSchemaSnapshotReader
    {
        public static ContentTypeSchemaSnapshot Read(
            ClientContext context,
            string contentTypeId,
            IEnumerable<string> requiredFieldIdentifiers,
            ICollection<string> diagnostics)
        {
            return Read(context, context == null ? null : context.Web, contentTypeId, requiredFieldIdentifiers, diagnostics);
        }

        public static ContentTypeSchemaSnapshot Read(
            ClientContext context,
            Web web,
            string contentTypeId,
            IEnumerable<string> requiredFieldIdentifiers,
            ICollection<string> diagnostics)
        {
            if (context == null)
            {
                throw new ArgumentNullException(nameof(context));
            }

            if (string.IsNullOrWhiteSpace(contentTypeId))
            {
                return Missing("The associated content type ID is unavailable.");
            }

            if (web == null)
            {
                throw new ArgumentNullException(nameof(web));
            }

            // The Page Layout association already gives us the exact content type ID.
            // Avoid enumerating the whole inherited content type catalog while retaining
            // the complete AvailableFields enumeration used for field-conflict evidence.
            var contentType = web.AvailableContentTypes.GetById(contentTypeId);
            var fields = web.AvailableFields;
            context.Load(web, value => value.Url);
            context.Load(contentType,
                value => value.Id,
                value => value.Name,
                value => value.Description,
                value => value.Group,
                value => value.Scope,
                value => value.ReadOnly,
                value => value.Sealed,
                value => value.Hidden,
                value => value.Parent);
            context.Load(fields, values => values.Include(
                value => value.Id,
                value => value.InternalName,
                value => value.Title,
                value => value.TypeAsString,
                value => value.Group,
                value => value.Required,
                value => value.Hidden,
                value => value.ReadOnlyField,
                value => value.Sealed,
                value => value.SchemaXml));
            context.ExecuteQueryRetry();
            if (contentType.ServerObjectIsNull.GetValueOrDefault(true))
            {
                return Missing($"Associated content type '{contentTypeId}' was not found in the source web.");
            }

            context.Load(contentType.FieldLinks, values => values.Include(
                value => value.Id,
                value => value.Name,
                value => value.Required,
                value => value.Hidden));
            if (contentType.Parent != null)
            {
                context.Load(contentType.Parent, value => value.Id, value => value.Name);
                context.Load(contentType.Parent.FieldLinks, values => values.Include(value => value.Id));
            }

            context.ExecuteQueryRetry();
            var requiredIdentifiers = new HashSet<string>(
                (requiredFieldIdentifiers ?? Enumerable.Empty<string>())
                    .Select(NormalizeFieldIdentifier)
                    .Where(value => !string.IsNullOrWhiteSpace(value)),
                StringComparer.OrdinalIgnoreCase);
            var parentFieldIds = contentType.Parent == null
                ? new HashSet<Guid>()
                : new HashSet<Guid>(contentType.Parent.FieldLinks.Select(value => value.Id));
            var requiredLinks = contentType.FieldLinks
                .Where(link => MatchesFieldIdentifier(requiredIdentifiers, link.Id, link.Name))
                .Select(link => new ContentTypeFieldLinkSnapshot
                {
                    FieldId = link.Id,
                    Name = link.Name,
                    Required = link.Required,
                    Hidden = link.Hidden,
                    Role = parentFieldIds.Contains(link.Id) ? FieldSchemaRole.InheritedFromParent : FieldSchemaRole.DirectBinding
                })
                .OrderBy(link => link.Name, StringComparer.OrdinalIgnoreCase)
                .ToList();
            var matchedIdentifiers = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (var link in requiredLinks)
            {
                matchedIdentifiers.Add(NormalizeFieldIdentifier(link.Name));
                matchedIdentifiers.Add(NormalizeFieldIdentifier(link.FieldId.ToString("D")));
            }
            var missingIdentifiers = requiredIdentifiers.Except(matchedIdentifiers, StringComparer.OrdinalIgnoreCase).ToArray();
            foreach (var missingIdentifier in missingIdentifiers)
            {
                diagnostics?.Add($"Associated content type '{contentType.Name}' does not expose required field link identifier '{missingIdentifier}'.");
            }

            var fieldGroups = fields
                .Where(value => value != null)
                .GroupBy(value => value.Id)
                .ToArray();
            var hasConflictingDuplicateFields = false;
            foreach (var group in fieldGroups.Where(value => value.Count() > 1))
            {
                var distinctSchemas = group
                    .Select(value => value.SchemaXml ?? string.Empty)
                    .Distinct(StringComparer.Ordinal)
                    .Count();
                if (distinctSchemas > 1)
                {
                    hasConflictingDuplicateFields = true;
                    diagnostics?.Add($"Available field ID '{group.Key:D}' returned {group.Count()} rows with conflicting schema evidence; the lexically first schema is retained and the content type evidence is partial.");
                }
                else
                {
                    diagnostics?.Add($"Available field ID '{group.Key:D}' returned {group.Count()} identical rows; duplicate enumeration evidence was collapsed.");
                }
            }
            var fieldsById = fieldGroups.ToDictionary(
                group => group.Key,
                group => group
                    .OrderBy(value => value.SchemaXml ?? string.Empty, StringComparer.Ordinal)
                    .ThenBy(value => value.InternalName ?? string.Empty, StringComparer.OrdinalIgnoreCase)
                    .First());
            var closure = new List<FieldSchemaSnapshot>();
            foreach (var link in requiredLinks)
            {
                Field sourceField;
                if (!fieldsById.TryGetValue(link.FieldId, out sourceField))
                {
                    diagnostics?.Add($"Required field schema '{link.Name}' ({link.FieldId:D}) was not found in the source web.");
                    continue;
                }

                var snapshot = CreateFieldSnapshot(context, sourceField, link.Role);
                var existingIndex = closure.FindIndex(value => value.Id == snapshot.Id);
                if (existingIndex < 0)
                {
                    closure.Add(snapshot);
                }
                else if (closure[existingIndex].Role == FieldSchemaRole.Dependency
                    && snapshot.Role != FieldSchemaRole.Dependency)
                {
                    // A taxonomy hidden-text dependency can also be an explicit Content Type
                    // field link. Keep one logical field node and promote its direct role.
                    closure[existingIndex] = snapshot;
                }
                if (snapshot.Taxonomy != null
                    && snapshot.Taxonomy.HiddenTextFieldId != Guid.Empty
                    && fieldsById.TryGetValue(snapshot.Taxonomy.HiddenTextFieldId, out var hiddenField)
                    && closure.All(value => value.Id != hiddenField.Id))
                {
                    closure.Add(CreateFieldSnapshot(context, hiddenField, FieldSchemaRole.Dependency));
                }
            }

            foreach (var field in closure)
            {
                field.Ownership = FieldOwnershipClassifier.Classify(field, closure);
            }

            var complete = missingIdentifiers.Length == 0
                && !hasConflictingDuplicateFields
                && requiredLinks.Count == closure.Count(value => value.Role != FieldSchemaRole.Dependency);
            return new ContentTypeSchemaSnapshot
            {
                EvidenceState = complete ? ContentTypeSchemaEvidenceState.Readable : ContentTypeSchemaEvidenceState.Partial,
                SourceWebUrl = web.Url,
                SourceScope = contentType.Scope,
                ContentTypeId = contentType.Id.StringValue,
                Name = contentType.Name,
                Description = contentType.Description,
                Group = contentType.Group,
                ReadOnly = contentType.ReadOnly,
                Sealed = contentType.Sealed,
                Hidden = contentType.Hidden,
                ParentContentTypeId = contentType.Parent?.Id.StringValue,
                ParentContentTypeName = contentType.Parent?.Name,
                RequiredFieldLinks = requiredLinks,
                RequiredFieldClosure = closure.OrderBy(value => value.Role == FieldSchemaRole.Dependency ? 0 : 1).ThenBy(value => value.InternalName, StringComparer.OrdinalIgnoreCase).ToList(),
                Availability = complete ? EvidenceAvailability.Captured : EvidenceAvailability.Partial,
                Diagnostics = diagnostics == null ? new List<string>() : diagnostics.ToList()
            };
        }

        private static FieldSchemaSnapshot CreateFieldSnapshot(
            ClientContext context,
            Field field,
            FieldSchemaRole role)
        {
            var snapshot = new FieldSchemaSnapshot
            {
                Id = field.Id,
                InternalName = field.InternalName,
                Title = field.Title,
                TypeAsString = field.TypeAsString,
                Group = field.Group,
                Required = field.Required,
                Hidden = field.Hidden,
                ReadOnly = field.ReadOnlyField,
                Sealed = field.Sealed,
                SchemaXml = field.SchemaXml,
                SchemaXmlSha256 = MigrationDigest.ComputeSha256(field.SchemaXml ?? string.Empty),
                PortableSchemaSha256 = FieldSchemaCanonicalizer.PortableDigest(field.SchemaXml),
                Role = role
            };
            if (field.TypeAsString.StartsWith("TaxonomyFieldType", StringComparison.OrdinalIgnoreCase))
            {
                var taxonomyField = context.CastTo<TaxonomyField>(field);
                context.Load(taxonomyField,
                    value => value.SspId,
                    value => value.TermSetId,
                    value => value.AnchorId,
                    value => value.TextField,
                    value => value.Open);
                context.ExecuteQueryRetry();
                snapshot.Taxonomy = new TaxonomyFieldBindingSnapshot
                {
                    SourceTermStoreId = taxonomyField.SspId,
                    SourceTermSetId = taxonomyField.TermSetId,
                    AnchorTermId = taxonomyField.AnchorId,
                    HiddenTextFieldId = taxonomyField.TextField,
                    Open = taxonomyField.Open
                };
            }

            return snapshot;
        }

        private static ContentTypeSchemaSnapshot Missing(string diagnostic)
        {
            return new ContentTypeSchemaSnapshot
            {
                EvidenceState = ContentTypeSchemaEvidenceState.Missing,
                Availability = EvidenceAvailability.Unavailable,
                Diagnostics = new List<string> { diagnostic }
            };
        }

        internal static bool MatchesFieldIdentifier(
            IEnumerable<string> requiredFieldIdentifiers,
            Guid fieldId,
            string fieldName)
        {
            var identifiers = new HashSet<string>(
                (requiredFieldIdentifiers ?? Enumerable.Empty<string>())
                    .Select(NormalizeFieldIdentifier)
                    .Where(value => !string.IsNullOrWhiteSpace(value)),
                StringComparer.OrdinalIgnoreCase);
            return MatchesFieldIdentifier(identifiers, fieldId, fieldName);
        }

        private static bool MatchesFieldIdentifier(
            ISet<string> requiredFieldIdentifiers,
            Guid fieldId,
            string fieldName)
        {
            return requiredFieldIdentifiers.Contains(NormalizeFieldIdentifier(fieldName))
                || requiredFieldIdentifiers.Contains(NormalizeFieldIdentifier(fieldId.ToString("D")));
        }

        private static string NormalizeFieldIdentifier(string value)
        {
            return (value ?? string.Empty).Trim().Trim('{', '}');
        }
    }
}
