using Microsoft.SharePoint.Client;
using PnP.Framework.Migration.Schema.Fields;
using System;
using System.Collections.Generic;
using System.Linq;

namespace PnP.Framework.Migration.Schema.ContentTypes
{
    internal static class ContentTypeMaterializer
    {
        public static ContentTypeMaterializationDisposition Ensure(
            ClientContext context,
            Web web,
            ContentTypeMaterializationPlan plan,
            ContentTypeTargetAdmission admission)
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

            if (admission == null || !admission.IsEligible || admission.Disposition == ContentTypeMaterializationDisposition.Block)
            {
                throw new InvalidOperationException("Blocked content type schema cannot be materialized.");
            }

            if (plan.Disposition == ContentTypeMaterializationDisposition.ReuseOwned
                && admission.Disposition != ContentTypeMaterializationDisposition.ReuseOwned)
            {
                throw new InvalidOperationException("A target-runtime-only content type plan cannot create or repair schema.");
            }

            if (admission.Disposition == ContentTypeMaterializationDisposition.ReuseOwned)
            {
                Verify(context, web, plan);
                return ContentTypeMaterializationDisposition.ReuseOwned;
            }

            var preflight = ContentTypeTargetInspector.Inspect(context, web, plan);
            var preflightAdmission = ContentTypeTargetAdmissionEvaluator.Evaluate(plan, preflight);
            if (!preflightAdmission.IsEligible)
            {
                throw new InvalidOperationException("Fresh content type preflight no longer permits create-only materialization.");
            }

            if (preflightAdmission.Disposition == ContentTypeMaterializationDisposition.ReuseOwned)
            {
                Verify(context, web, plan);
                return ContentTypeMaterializationDisposition.ReuseOwned;
            }

            if (preflightAdmission.Disposition != ContentTypeMaterializationDisposition.CreateOwned)
            {
                throw new InvalidOperationException($"Unexpected content type schema admission disposition: {preflightAdmission.Disposition}.");
            }

            var probedFields = preflight.Fields.ToDictionary(value => value.FieldId);
            foreach (var fieldPlan in plan.Fields
                         .Where(value => value.Disposition == FieldSchemaMaterializationDisposition.CreateOrReuseOwned)
                         .OrderBy(value => value.Role == FieldSchemaRole.Dependency ? 0 : 1)
                         .ThenBy(value => value.FieldId))
            {
                var probe = probedFields[fieldPlan.FieldId];
                if (probe.Exists)
                {
                    continue;
                }

                if (string.IsNullOrWhiteSpace(fieldPlan.TargetSchemaXml))
                {
                    throw new InvalidOperationException($"Target field schema is unavailable: {fieldPlan.InternalName} ({fieldPlan.FieldId:D}).");
                }

                var created = web.Fields.AddFieldAsXml(
                    fieldPlan.TargetSchemaXml,
                    false,
                    AddFieldOptions.AddFieldInternalNameHint);
                context.Load(created, value => value.Id, value => value.InternalName, value => value.TypeAsString, value => value.SchemaXml);
                context.ExecuteQueryRetry();
                var actualDigest = FieldSchemaCanonicalizer.PortableDigest(created.SchemaXml);
                if (created.Id != fieldPlan.FieldId
                    || !string.Equals(actualDigest, fieldPlan.TargetPortableSchemaSha256, StringComparison.OrdinalIgnoreCase))
                {
                    throw new InvalidOperationException($"Fresh field readback differs from the sealed plan: {fieldPlan.InternalName} ({fieldPlan.FieldId:D}).");
                }
            }

            var parent = web.AvailableContentTypes.GetById(plan.ParentContentTypeId);
            context.Load(parent, value => value.Id, value => value.Name);
            context.ExecuteQueryRetry();
            if (!string.Equals(parent.Id.StringValue, plan.ParentContentTypeId, StringComparison.OrdinalIgnoreCase))
            {
                throw new InvalidOperationException($"Target parent content type is unavailable: {plan.ParentContentTypeId}.");
            }

            var createdContentType = web.ContentTypes.Add(new ContentTypeCreationInformation
            {
                Id = plan.ContentTypeId,
                Name = plan.Name,
                Description = plan.Description,
                Group = plan.Group
            });
            context.Load(createdContentType, value => value.Id, value => value.Name);
            context.ExecuteQueryRetry();
            if (!string.Equals(createdContentType.Id.StringValue, plan.ContentTypeId, StringComparison.OrdinalIgnoreCase))
            {
                throw new InvalidOperationException(
                    $"SharePoint created content type ID '{createdContentType.Id.StringValue}' instead of sealed ID '{plan.ContentTypeId}'.");
            }

            context.Load(createdContentType.FieldLinks, values => values.Include(
                value => value.Id,
                value => value.Required,
                value => value.Hidden));
            context.ExecuteQueryRetry();
            var existingLinks = createdContentType.FieldLinks.ToDictionary(value => value.Id);
            foreach (var linkPlan in plan.RequiredFieldLinks
                         .Where(value => value.Role != FieldSchemaRole.InheritedFromParent)
                         .OrderBy(value => value.Role == FieldSchemaRole.Dependency ? 0 : 1)
                         .ThenBy(value => value.FieldId))
            {
                FieldLink link;
                if (!existingLinks.TryGetValue(linkPlan.FieldId, out link))
                {
                    var field = web.Fields.GetById(linkPlan.FieldId);
                    link = createdContentType.FieldLinks.Add(new FieldLinkCreationInformation { Field = field });
                    existingLinks[linkPlan.FieldId] = link;
                }
                link.Required = linkPlan.Required;
                link.Hidden = linkPlan.Hidden;
            }

            createdContentType.Hidden = plan.Hidden;
            createdContentType.ReadOnly = plan.ReadOnly;
            createdContentType.Sealed = plan.Sealed;
            createdContentType.Update(false);
            context.ExecuteQueryRetry();
            Verify(context, web, plan);
            return ContentTypeMaterializationDisposition.CreateOwned;
        }

        private static ContentTypeTargetProbe Verify(
            ClientContext context,
            Web web,
            ContentTypeMaterializationPlan plan)
        {
            var readback = ContentTypeTargetInspector.Inspect(context, web, plan);
            var readbackAdmission = ContentTypeTargetAdmissionEvaluator.Evaluate(plan, readback);
            if (!readbackAdmission.IsEligible
                || readbackAdmission.Disposition != ContentTypeMaterializationDisposition.ReuseOwned)
            {
                throw new InvalidOperationException("Fresh content type schema readback differs from the sealed plan.");
            }

            return readback;
        }
    }
}
