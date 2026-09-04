using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.Taxonomy;
using PnP.Framework.Migration.Evidence;
using PnP.Framework.Migration.Schema.Fields;
using PnP.Framework.Migration.Taxonomy;
using System;
using System.Collections.Generic;
using System.Linq;

namespace PnP.Framework.Migration.Schema.ContentTypes
{
    internal static class ContentTypeTargetInspector
    {
        public static ContentTypeTargetProbe Inspect(
            ClientContext context,
            Web web,
            ContentTypeMaterializationPlan plan)
        {
            if (context == null)
            {
                throw new ArgumentNullException(nameof(context));
            }

            if (web == null)
            {
                throw new ArgumentNullException(nameof(web));
            }

            if (plan == null)
            {
                throw new ArgumentNullException(nameof(plan));
            }

            var diagnostics = new List<string>();
            var siteContentTypes = web.ContentTypes;
            var availableContentTypes = web.AvailableContentTypes;
            var fields = web.AvailableFields;
            context.Load(web, value => value.EffectiveBasePermissions);
            context.Load(siteContentTypes, values => values.Include(
                value => value.Id,
                value => value.Name));
            context.Load(availableContentTypes, values => values.Include(
                value => value.Id,
                value => value.Name));
            context.Load(fields, values => values.Include(
                value => value.Id,
                value => value.InternalName,
                value => value.Title,
                value => value.TypeAsString,
                value => value.SchemaXml));
            context.ExecuteQueryRetry();

            // A missing target Content Type is a normal CreateMissing observation.
            // CSOM GetById returns a null server object whose nested Parent or
            // FieldLinks cannot be loaded safely, so resolve candidates from the
            // loaded collections and request details only for objects that exist.
            var parent = availableContentTypes.FirstOrDefault(value =>
                string.Equals(value.Id.StringValue, plan.ParentContentTypeId, StringComparison.OrdinalIgnoreCase));
            var exact = siteContentTypes.FirstOrDefault(value =>
                string.Equals(value.Id.StringValue, plan.ContentTypeId, StringComparison.OrdinalIgnoreCase));
            if (parent != null)
            {
                context.Load(parent.FieldLinks, values => values.Include(
                    value => value.Id,
                    value => value.Name,
                    value => value.Required,
                    value => value.Hidden));
            }
            if (exact != null)
            {
                context.Load(exact,
                    value => value.Id,
                    value => value.Name,
                    value => value.Description,
                    value => value.Group,
                    value => value.ReadOnly,
                    value => value.Sealed,
                    value => value.Hidden);
                context.Load(exact.Parent, value => value.Id);
                context.Load(exact.FieldLinks, values => values.Include(
                    value => value.Id,
                    value => value.Name,
                    value => value.Required,
                    value => value.Hidden));
            }
            if (parent != null || exact != null)
            {
                context.ExecuteQueryRetry();
            }

            var unresolvedTermSetProbes = new Dictionary<Guid, (bool Exists, string Name)>();
            var unresolvedGroups = plan.Fields
                .Where(value => value.TaxonomyMappingMode == TaxonomyTargetMappingMode.PreserveUnresolvedSourceReference)
                .Where(value => value.TargetTermStoreId.HasValue && value.TargetTermSetId.HasValue)
                .GroupBy(value => value.TargetTermStoreId.Value.ToString("D") + "/" + value.TargetTermSetId.Value.ToString("D"), StringComparer.Ordinal)
                .ToArray();
            if (unresolvedGroups.Length > 0)
            {
                var taxonomySession = TaxonomySession.GetTaxonomySession(context);
                var pending = unresolvedGroups.Select(group =>
                {
                    var first = group.First();
                    var store = taxonomySession.TermStores.GetById(first.TargetTermStoreId.Value);
                    var termSet = store.GetTermSet(first.TargetTermSetId.Value);
                    context.Load(termSet, value => value.Id, value => value.Name);
                    return new { Fields = group.ToArray(), TermSet = termSet };
                }).ToArray();
                context.ExecuteQueryRetry();
                foreach (var item in pending)
                {
                    var exists = !item.TermSet.ServerObjectIsNull.GetValueOrDefault(true);
                    foreach (var field in item.Fields)
                    {
                        unresolvedTermSetProbes[field.FieldId] = (exists, exists ? item.TermSet.Name : null);
                    }
                }
            }

            var fieldById = fields.ToDictionary(value => value.Id);
            var fieldProbes = new List<FieldSchemaTargetProbe>();
            foreach (var desired in plan.Fields.OrderBy(value => value.FieldId))
            {
                Field field;
                fieldById.TryGetValue(desired.FieldId, out field);
                string portableDigest = null;
                if (field != null && !string.IsNullOrWhiteSpace(field.SchemaXml))
                {
                    try
                    {
                        portableDigest = FieldSchemaCanonicalizer.PortableDigest(field.SchemaXml);
                    }
                    catch (Exception exception) when (exception is ArgumentException || exception is System.Xml.XmlException)
                    {
                        diagnostics.Add($"Target field '{field.InternalName}' ({field.Id:D}) has invalid SchemaXml: {exception.Message}");
                    }
                }
                var hasUnresolvedProbe = unresolvedTermSetProbes.TryGetValue(desired.FieldId, out var unresolvedProbe);

                fieldProbes.Add(new FieldSchemaTargetProbe
                {
                    FieldId = desired.FieldId,
                    Exists = field != null,
                    InternalName = field?.InternalName,
                    Title = field?.Title,
                    TypeAsString = field?.TypeAsString,
                    PortableSchemaSha256 = portableDigest,
                    UnresolvedTargetTermSetExists = desired.TaxonomyMappingMode == TaxonomyTargetMappingMode.PreserveUnresolvedSourceReference
                        ? hasUnresolvedProbe ? unresolvedProbe.Exists : (bool?)null
                        : (bool?)null,
                    UnresolvedTargetTermSetName = desired.TaxonomyMappingMode == TaxonomyTargetMappingMode.PreserveUnresolvedSourceReference
                        ? unresolvedProbe.Name
                        : null
                });
            }

            return new ContentTypeTargetProbe
            {
                ContentTypeId = plan.ContentTypeId,
                ParentContentTypeAvailable = parent != null,
                ResolvedParentContentTypeId = parent?.Id.StringValue,
                ParentFieldLinks = parent == null ? new List<ContentTypeFieldLinkTargetProbe>() : Links(parent.FieldLinks),
                ContentTypeExists = exact != null,
                ExistingName = exact?.Name,
                ExistingDescription = exact?.Description,
                ExistingGroup = exact?.Group,
                ExistingReadOnly = exact != null && exact.ReadOnly,
                ExistingSealed = exact != null && exact.Sealed,
                ExistingHidden = exact != null && exact.Hidden,
                ExistingParentContentTypeId = exact?.Parent?.Id.StringValue,
                ExistingFieldLinks = exact == null ? new List<ContentTypeFieldLinkTargetProbe>() : Links(exact.FieldLinks),
                SameNameDifferentIds = siteContentTypes
                    .Where(value => string.Equals(value.Name, plan.Name, StringComparison.OrdinalIgnoreCase))
                    .Where(value => !string.Equals(value.Id.StringValue, plan.ContentTypeId, StringComparison.OrdinalIgnoreCase))
                    .Select(value => value.Id.StringValue)
                    .OrderBy(value => value, StringComparer.Ordinal)
                    .ToList(),
                Fields = fieldProbes,
                CanManageContentTypes = web.EffectiveBasePermissions.Has(PermissionKind.ManageLists),
                Availability = EvidenceAvailability.Captured,
                Diagnostics = diagnostics
            };
        }

        private static IList<ContentTypeFieldLinkTargetProbe> Links(FieldLinkCollection links)
        {
            return links.Select(value => new ContentTypeFieldLinkTargetProbe
                {
                    FieldId = value.Id,
                    Name = value.Name,
                    Required = value.Required,
                    Hidden = value.Hidden
                })
                .OrderBy(value => value.FieldId)
                .ToList();
        }
    }
}
