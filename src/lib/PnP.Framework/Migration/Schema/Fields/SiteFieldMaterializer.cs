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

            var merged = (plans ?? Enumerable.Empty<FieldSchemaMaterializationPlan>())
                .Where(value => value != null)
                .GroupBy(value => value.FieldId)
                .Select(group => Merge(group.Key, group))
                .ToArray();
            var values = OrderForMaterialization(merged)
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
            context.Load(web.Fields, fields => fields.Include(
                value => value.Id,
                value => value.InternalName,
                value => value.Title,
                value => value.TypeAsString,
                value => value.SchemaXml));
            context.ExecuteQueryRetry();

            var byId = web.Fields.AsEnumerable().ToDictionary(value => value.Id);
            var byName = web.Fields.AsEnumerable()
                .Where(value => !string.IsNullOrWhiteSpace(value.InternalName))
                .GroupBy(value => value.InternalName, StringComparer.OrdinalIgnoreCase)
                .ToDictionary(group => group.Key, group => group.ToArray(), StringComparer.OrdinalIgnoreCase);
            var createdCount = 0;
            foreach (var plan in values)
            {
                byId.TryGetValue(plan.FieldId, out var existing);
                if (existing != null)
                {
                    EnsureDisplayName(context, existing, plan);
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
                if (plan.Disposition == FieldSchemaMaterializationDisposition.RequireTargetRuntime
                    && string.IsNullOrWhiteSpace(plan.TargetSchemaXml))
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
                    value => value.Title,
                    value => value.TypeAsString,
                    value => value.SchemaXml);
                context.ExecuteQueryRetry();
                Verify(created, plan);
                byId[created.Id] = created;
                if (!byName.TryGetValue(created.InternalName, out var sameName))
                {
                    byName[created.InternalName] = new[] { created };
                }
                else
                {
                    byName[created.InternalName] = sameName.Concat(new[] { created }).ToArray();
                }
                createdCount++;
            }

            // Re-read every selected identity after all dependent/calculated
            // fields have been created; this is the transaction's fresh proof.
            foreach (var plan in values)
            {
                var readback = web.Fields.GetById(plan.FieldId);
                context.Load(readback,
                    value => value.Id,
                    value => value.InternalName,
                    value => value.Title,
                    value => value.TypeAsString,
                    value => value.SchemaXml);
                context.ExecuteQueryRetry();
                Verify(readback, plan);
            }
            return createdCount;
        }

        private static void EnsureDisplayName(
            ClientContext context,
            Field field,
            FieldSchemaMaterializationPlan plan)
        {
            if (!string.IsNullOrWhiteSpace(field.Title) || string.IsNullOrWhiteSpace(plan.Title))
            {
                return;
            }

            field.Title = plan.Title;
            field.Update();
            context.ExecuteQueryRetry();
            context.Load(field,
                value => value.Id,
                value => value.InternalName,
                value => value.Title,
                value => value.TypeAsString,
                value => value.SchemaXml);
            context.ExecuteQueryRetry();
        }

        internal static IEnumerable<FieldSchemaMaterializationPlan> OrderForMaterialization(
            IEnumerable<FieldSchemaMaterializationPlan> plans)
        {
            var values = (plans ?? Enumerable.Empty<FieldSchemaMaterializationPlan>()).ToArray();
            var hiddenTextFieldIds = values
                .Where(value => value?.HiddenTextFieldId.HasValue == true)
                .Select(value => value.HiddenTextFieldId.Value)
                .ToHashSet();
            return values
                .OrderBy(value => hiddenTextFieldIds.Contains(value.FieldId)
                    ? 0
                    : value.HiddenTextFieldId.HasValue ? 2 : 1)
                .ThenBy(value => value.TypeAsString?.StartsWith("Calculated", StringComparison.OrdinalIgnoreCase) == true ? 1 : 0)
                .ThenBy(value => value.Role == FieldSchemaRole.Dependency ? 0 : 1)
                .ThenBy(value => value.FieldId);
        }

        internal static FieldSchemaMaterializationPlan Merge(
            Guid fieldId,
            IEnumerable<FieldSchemaMaterializationPlan> candidates)
        {
            var values = candidates.ToArray();
            var producers = values
                .Where(value => value.Disposition == FieldSchemaMaterializationDisposition.CreateOrReuseOwned)
                .OrderBy(value => value.Role == FieldSchemaRole.DirectBinding ? 0 : 1)
                .ToArray();
            if (producers.Length > 0)
            {
                var producer = producers[0];
                if (producers.Any(value => !EquivalentProducer(value, producer))
                    || values.Where(value => value.Disposition == FieldSchemaMaterializationDisposition.RequireTargetRuntime)
                        .Any(value => !EquivalentInheritedConsumer(value, producer))
                    || values.Any(value => value.Disposition == FieldSchemaMaterializationDisposition.Block))
                {
                    throw Conflict(fieldId);
                }
                return producer;
            }

            var first = values[0];
            if (values.Any(value => value.Disposition != first.Disposition
                || value.Ownership != first.Ownership
                || !EquivalentIdentity(value, first)
                || !string.Equals(
                    value.SourcePortableSchemaSha256 ?? string.Empty,
                    first.SourcePortableSchemaSha256 ?? string.Empty,
                    StringComparison.OrdinalIgnoreCase)))
            {
                throw Conflict(fieldId);
            }
            return first;
        }

        private static bool EquivalentProducer(
            FieldSchemaMaterializationPlan left,
            FieldSchemaMaterializationPlan right) =>
            left.Ownership == right.Ownership
            && EquivalentIdentity(left, right)
            && EquivalentTargetSchema(left, right);

        private static bool EquivalentInheritedConsumer(
            FieldSchemaMaterializationPlan consumer,
            FieldSchemaMaterializationPlan producer) =>
            consumer.Role == FieldSchemaRole.InheritedFromParent
            && consumer.Ownership == FieldOwnership.TargetRuntime
            && EquivalentIdentity(consumer, producer)
            && string.Equals(
                consumer.SourcePortableSchemaSha256 ?? string.Empty,
                producer.SourcePortableSchemaSha256 ?? string.Empty,
                StringComparison.OrdinalIgnoreCase);

        private static bool EquivalentIdentity(
            FieldSchemaMaterializationPlan left,
            FieldSchemaMaterializationPlan right) =>
            left.FieldId == right.FieldId
            && string.Equals(left.InternalName, right.InternalName, StringComparison.OrdinalIgnoreCase)
            && string.Equals(left.TypeAsString, right.TypeAsString, StringComparison.OrdinalIgnoreCase);

        private static InvalidDataException Conflict(Guid fieldId) =>
            new InvalidDataException(
                "Conflicting site-field execution plans were sealed for field "
                + fieldId.ToString("D") + ".");

        private static bool EquivalentTargetSchema(
            FieldSchemaMaterializationPlan left,
            FieldSchemaMaterializationPlan right)
        {
            if (string.IsNullOrWhiteSpace(left.TargetSchemaXml)
                || string.IsNullOrWhiteSpace(right.TargetSchemaXml))
            {
                return string.Equals(
                    left.TargetSchemaXml ?? string.Empty,
                    right.TargetSchemaXml ?? string.Empty,
                    StringComparison.Ordinal);
            }

            return string.Equals(
                FieldSchemaCanonicalizer.PortableDigest(left.TargetSchemaXml),
                FieldSchemaCanonicalizer.PortableDigest(right.TargetSchemaXml),
                StringComparison.OrdinalIgnoreCase);
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
            if (plan.Disposition == FieldSchemaMaterializationDisposition.CreateOrReuseOwned
                && !string.IsNullOrWhiteSpace(plan.Title)
                && string.IsNullOrWhiteSpace(field.Title))
            {
                throw new InvalidDataException(
                    "Fresh target site-field DisplayName is empty for '"
                    + plan.InternalName + "' (" + plan.FieldId.ToString("D") + ").");
            }
            if (plan.Disposition == FieldSchemaMaterializationDisposition.CreateOrReuseOwned
                || plan.Disposition == FieldSchemaMaterializationDisposition.RequireTargetRuntime
                    && !string.IsNullOrWhiteSpace(plan.TargetSchemaXml))
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
