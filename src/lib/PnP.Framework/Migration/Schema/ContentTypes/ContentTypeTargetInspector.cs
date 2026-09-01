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
            var availableContentTypes = web.AvailableContentTypes;
            var fields = web.AvailableFields;
            context.Load(web, value => value.EffectiveBasePermissions);
            context.Load(siteContentTypes, values => values.Include(
                value => value.Id,
                value => value.Name,
                value => value.Description,
                value => value.Group));
            context.Load(availableContentTypes, values => values.Include(value => value.Id, value => value.Name));
            context.Load(fields, values => values.Include(
                value => value.Id,
                value => value.InternalName,
                value => value.Title,
                value => value.TypeAsString,
                value => value.SchemaXml));
            context.ExecuteQueryRetry();

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
