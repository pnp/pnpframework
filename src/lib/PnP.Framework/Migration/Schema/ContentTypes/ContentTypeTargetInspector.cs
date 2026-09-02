using Microsoft.SharePoint.Client;
using PnP.Framework.Migration.Evidence;
using PnP.Framework.Migration.Schema.Fields;
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
            var fields = web.AvailableFields;
            var parentCandidate = web.AvailableContentTypes.GetById(plan.ParentContentTypeId);
            var exactCandidate = siteContentTypes.GetById(plan.ContentTypeId);
            context.Load(web, value => value.EffectiveBasePermissions);
            context.Load(siteContentTypes, values => values.Include(
                value => value.Id,
                value => value.Name));
            context.Load(fields, values => values.Include(
                value => value.Id,
                value => value.InternalName,
                value => value.Title,
                value => value.TypeAsString,
                value => value.SchemaXml));
            context.Load(parentCandidate, value => value.Id);
            context.Load(parentCandidate.FieldLinks, values => values.Include(
                value => value.Id,
                value => value.Name,
                value => value.Required,
                value => value.Hidden));
            context.Load(exactCandidate,
                value => value.Id,
                value => value.Name,
                value => value.Description,
                value => value.Group,
                value => value.ReadOnly,
                value => value.Sealed,
                value => value.Hidden);
            context.Load(exactCandidate.Parent, value => value.Id);
            context.Load(exactCandidate.FieldLinks, values => values.Include(
                value => value.Id,
                value => value.Name,
                value => value.Required,
                value => value.Hidden));
            context.ExecuteQueryRetry();

            var parent = parentCandidate.ServerObjectIsNull.GetValueOrDefault(true)
                || !string.Equals(parentCandidate.Id.StringValue, plan.ParentContentTypeId, StringComparison.OrdinalIgnoreCase)
                ? null
                : parentCandidate;
            var exact = exactCandidate.ServerObjectIsNull.GetValueOrDefault(true)
                || !string.Equals(exactCandidate.Id.StringValue, plan.ContentTypeId, StringComparison.OrdinalIgnoreCase)
                ? null
                : exactCandidate;

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

                fieldProbes.Add(new FieldSchemaTargetProbe
                {
                    FieldId = desired.FieldId,
                    Exists = field != null,
                    InternalName = field?.InternalName,
                    Title = field?.Title,
                    TypeAsString = field?.TypeAsString,
                    PortableSchemaSha256 = portableDigest
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
