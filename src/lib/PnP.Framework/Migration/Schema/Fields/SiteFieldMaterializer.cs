using Microsoft.SharePoint.Client;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace PnP.Framework.Migration.Schema.Fields
{
    /// <summary>
    /// Materializes an admitted subset of site-field schema without requiring
    /// the content type that may eventually consume those fields to be
    /// executable in the same run.
    /// </summary>
    internal static class SiteFieldMaterializer
    {
        public static int Ensure(
            ClientContext context,
            Web web,
            IEnumerable<FieldSchemaMaterializationPlan> plans)
        {
            if (context == null)
            {
                throw new ArgumentNullException(nameof(context));
            }
            if (web == null)
            {
                throw new ArgumentNullException(nameof(web));
            }

            var values = (plans ?? Enumerable.Empty<FieldSchemaMaterializationPlan>())
                .Where(value => value != null)
                .GroupBy(value => value.FieldId)
                .Select(group => Merge(group.Key, group))
                .OrderBy(value => value.TypeAsString?.StartsWith("Calculated", StringComparison.OrdinalIgnoreCase) == true ? 1 : 0)
                .ThenBy(value => value.Role == FieldSchemaRole.Dependency ? 0 : 1)
                .ThenBy(value => value.FieldId)
                .ToArray();
            if (values.Length == 0)
            {
                return 0;
            }
            if (values.Any(value => value.Disposition == FieldSchemaMaterializationDisposition.Block))
            {
                throw new InvalidOperationException("A blocked site-field schema cannot enter the execution scope.");
            }

            context.Load(web, value => value.EffectiveBasePermissions);
            context.Load(web.AvailableFields, fields => fields.Include(
                value => value.Id,
                value => value.InternalName,
                value => value.TypeAsString,
                value => value.SchemaXml));
            context.ExecuteQueryRetry();

            var byId = web.AvailableFields.AsEnumerable().ToDictionary(value => value.Id);
            var byName = web.AvailableFields.AsEnumerable()
                .Where(value => !string.IsNullOrWhiteSpace(value.InternalName))
                .GroupBy(value => value.InternalName, StringComparer.OrdinalIgnoreCase)
                .ToDictionary(group => group.Key, group => group.ToArray(), StringComparer.OrdinalIgnoreCase);
            var createdCount = 0;
            foreach (var plan in values)
            {
                byId.TryGetValue(plan.FieldId, out var existing);
                if (existing != null)
                {
                    Verify(existing, plan);
                    continue;
                }

                if (!string.IsNullOrWhiteSpace(plan.InternalName)
                    && byName.TryGetValue(plan.InternalName, out var collisions)
                    && collisions.Any(value => value.Id != plan.FieldId))
                {
                    throw new InvalidDataException(
                        "Target site-field name is already used by a different GUID: "
                        + plan.InternalName + ".");
                }
                if (plan.Disposition == FieldSchemaMaterializationDisposition.RequireTargetRuntime)
                {
                    throw new InvalidOperationException(
                        "The target runtime does not expose required site field '"
                        + plan.InternalName + "' (" + plan.FieldId.ToString("D") + ").");
                }
                if (!web.EffectiveBasePermissions.Has(PermissionKind.ManageLists))
                {
                    throw new UnauthorizedAccessException(
                        "The effective target permissions do not include ManageLists required to create site field '"
                        + plan.InternalName + "'.");
                }
                if (string.IsNullOrWhiteSpace(plan.TargetSchemaXml)
                    || string.IsNullOrWhiteSpace(plan.TargetPortableSchemaSha256))
                {
                    throw new InvalidDataException(
                        "The sealed target schema is unavailable for site field '"
                        + plan.InternalName + "' (" + plan.FieldId.ToString("D") + ").");
                }

                var created = web.Fields.AddFieldAsXml(
                    plan.TargetSchemaXml,
                    false,
                    AddFieldOptions.AddFieldInternalNameHint);
                context.Load(created,
                    value => value.Id,
                    value => value.InternalName,
                    value => value.TypeAsString,
                    value => value.SchemaXml);
                context.ExecuteQueryRetry();
                Verify(created, plan);
                createdCount++;
            }

            // Re-read every selected identity after all dependent/calculated
            // fields have been created; this is the transaction's fresh proof.
            foreach (var plan in values)
            {
                var readback = web.AvailableFields.GetById(plan.FieldId);
                context.Load(readback,
                    value => value.Id,
                    value => value.InternalName,
                    value => value.TypeAsString,
                    value => value.SchemaXml);
                context.ExecuteQueryRetry();
                Verify(readback, plan);
            }
            return createdCount;
        }

        private static FieldSchemaMaterializationPlan Merge(
            Guid fieldId,
            IEnumerable<FieldSchemaMaterializationPlan> candidates)
        {
            var values = candidates.ToArray();
            var first = values[0];
            if (values.Any(value =>
                value.Disposition != first.Disposition
                || value.Ownership != first.Ownership
                || !string.Equals(value.InternalName, first.InternalName, StringComparison.OrdinalIgnoreCase)
                || !string.Equals(value.TypeAsString, first.TypeAsString, StringComparison.OrdinalIgnoreCase)
                || !string.Equals(value.TargetSchemaXml ?? string.Empty, first.TargetSchemaXml ?? string.Empty, StringComparison.Ordinal)
                || !string.Equals(value.TargetPortableSchemaSha256 ?? string.Empty, first.TargetPortableSchemaSha256 ?? string.Empty, StringComparison.OrdinalIgnoreCase)))
            {
                throw new InvalidDataException(
                    "Conflicting site-field execution plans were sealed for field "
                    + fieldId.ToString("D") + ".");
            }
            return first;
        }

        private static void Verify(Field field, FieldSchemaMaterializationPlan plan)
        {
            var generatedCompanion = plan.Ownership == FieldOwnership.GeneratedTaxonomyCompanion;
            if (field == null
                || field.Id != plan.FieldId
                || (!generatedCompanion
                    && !string.Equals(field.InternalName, plan.InternalName, StringComparison.OrdinalIgnoreCase))
                || !string.Equals(field.TypeAsString, plan.TypeAsString, StringComparison.OrdinalIgnoreCase))
            {
                throw new InvalidDataException(
                    "Fresh target site-field identity/type differs for '"
                    + plan.InternalName + "' (" + plan.FieldId.ToString("D") + ").");
            }
            if (plan.Disposition == FieldSchemaMaterializationDisposition.CreateOrReuseOwned)
            {
                var digest = string.IsNullOrWhiteSpace(field.SchemaXml)
                    ? null
                    : FieldSchemaCanonicalizer.PortableDigest(field.SchemaXml);
                if (string.IsNullOrWhiteSpace(plan.TargetPortableSchemaSha256)
                    || !string.Equals(digest, plan.TargetPortableSchemaSha256, StringComparison.OrdinalIgnoreCase))
                {
                    throw new InvalidDataException(
                        "Fresh target site-field portable schema differs for '"
                        + plan.InternalName + "' (" + plan.FieldId.ToString("D") + ").");
                }
            }
        }
    }
}
