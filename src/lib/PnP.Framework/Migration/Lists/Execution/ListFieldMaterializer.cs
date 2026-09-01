using Microsoft.SharePoint.Client;
using PnP.Framework.Migration.Lists.Planning;
using PnP.Framework.Migration.Schema.Fields;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace PnP.Framework.Migration.Lists.Execution
{
    internal static class ListFieldMaterializer
    {
        public static void Ensure(
            ClientContext context,
            List targetList,
            ListMaterializationPlan plan,
            IDictionary<Guid, ListMaterializationReceipt> dependencyReceipts)
        {
            context.Load(targetList.Fields, values => values.Include(
                value => value.Id,
                value => value.InternalName,
                value => value.TypeAsString,
                value => value.SchemaXml));
            context.ExecuteQueryRetry();
            var targetById = targetList.Fields.AsEnumerable().ToDictionary(value => value.Id);
            foreach (var fieldPlan in plan.Fields.Where(value => value.Disposition != ListFieldMaterializationDisposition.EvidenceOnly
                && value.Disposition != ListFieldMaterializationDisposition.Block))
            {
                Field existing;
                targetById.TryGetValue(fieldPlan.SourceFieldId, out existing);
                if (fieldPlan.Disposition == ListFieldMaterializationDisposition.RequireTargetRuntime
                    || fieldPlan.Disposition == ListFieldMaterializationDisposition.RequireTargetRuntimeAndCopyValue)
                {
                    if (existing == null
                        || !string.Equals(existing.InternalName, fieldPlan.InternalName, StringComparison.OrdinalIgnoreCase)
                        || !ListFieldTypeCompatibility.IsCompatibleRuntimeType(existing.TypeAsString, fieldPlan.TypeAsString))
                    {
                        throw new InvalidDataException("Target List template does not expose required runtime field '" + fieldPlan.InternalName + "' (" + fieldPlan.SourceFieldId.ToString("D") + ").");
                    }
                    continue;
                }

                var targetSchema = fieldPlan.TargetSchemaXml;
                if (fieldPlan.Disposition == ListFieldMaterializationDisposition.MapLookup)
                {
                    ListMaterializationReceipt lookup;
                    if (!fieldPlan.SourceLookupListId.HasValue || !dependencyReceipts.TryGetValue(fieldPlan.SourceLookupListId.Value, out lookup))
                    {
                        throw new InvalidDataException("Lookup field '" + fieldPlan.InternalName + "' has no materialized dependency receipt.");
                    }
                    targetSchema = FieldSchemaCanonicalizer.RewriteLookupForTarget(fieldPlan.SourceSchemaXml, lookup.TargetWebId, lookup.TargetListId);
                }
                if (string.IsNullOrWhiteSpace(targetSchema))
                {
                    throw new InvalidDataException("Target field schema is unavailable for '" + fieldPlan.InternalName + "'.");
                }
                var expectedDigest = FieldSchemaCanonicalizer.PortableDigest(targetSchema);
                if (existing != null)
                {
                    if (!string.Equals(FieldSchemaCanonicalizer.PortableDigest(existing.SchemaXml), expectedDigest, StringComparison.OrdinalIgnoreCase))
                    {
                        throw new InvalidDataException("Target field GUID collision for '" + fieldPlan.InternalName + "' (" + fieldPlan.SourceFieldId.ToString("D") + ").");
                    }
                    continue;
                }

                var created = targetList.Fields.AddFieldAsXml(
                    targetSchema,
                    false,
                    AddFieldOptions.AddFieldInternalNameHint | AddFieldOptions.AddToNoContentType);
                context.Load(created, value => value.Id, value => value.InternalName, value => value.TypeAsString, value => value.SchemaXml);
                context.ExecuteQueryRetry();
                if (created.Id != fieldPlan.SourceFieldId
                    || !string.Equals(created.InternalName, fieldPlan.InternalName, StringComparison.OrdinalIgnoreCase)
                    || !string.Equals(FieldSchemaCanonicalizer.PortableDigest(created.SchemaXml), expectedDigest, StringComparison.OrdinalIgnoreCase))
                {
                    throw new InvalidDataException("Fresh target field readback differs from the sealed plan for '" + fieldPlan.InternalName + "'.");
                }
                targetById[created.Id] = created;
            }
        }

    }
}
