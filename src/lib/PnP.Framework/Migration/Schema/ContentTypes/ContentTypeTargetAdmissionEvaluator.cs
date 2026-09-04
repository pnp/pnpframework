using PnP.Framework.Migration.Diagnostics;
using PnP.Framework.Migration.Evidence;
using PnP.Framework.Migration.Schema.Fields;
using PnP.Framework.Migration.Taxonomy;
using System;
using System.Collections.Generic;
using System.Linq;

namespace PnP.Framework.Migration.Schema.ContentTypes
{
    public static class ContentTypeTargetAdmissionEvaluator
    {
        public static ContentTypeTargetAdmission Evaluate(
            ContentTypeMaterializationPlan plan,
            ContentTypeTargetProbe probe)
        {
            return Evaluate(plan, probe, null);
        }

        internal static ContentTypeTargetAdmission Evaluate(
            ContentTypeMaterializationPlan plan,
            ContentTypeTargetProbe probe,
            ContentTypeTargetAdmissionContext context)
        {
            var issues = new List<MigrationIssue>();
            var warnings = new List<string>();
            if (plan == null)
            {
                issues.Add(Issue("SourceContentTypeSchemaUnavailable", "source-content-type-schema",
                    "The source artifact has no readable required-field content type closure."));
                return Result(issues, warnings, ContentTypeMaterializationDisposition.Block);
            }

            if (probe == null || probe.Availability == EvidenceAvailability.Unavailable || probe.Availability == EvidenceAvailability.Conflict)
            {
                issues.Add(Issue("TargetContentTypeProbeUnavailable", "target-content-type-schema-probe",
                    $"The target content type probe is {(probe == null ? "missing" : probe.Availability.ToString())}."));
                return Result(issues, warnings, ContentTypeMaterializationDisposition.Block);
            }

            foreach (var field in plan.Fields.Where(value => value.Disposition == FieldSchemaMaterializationDisposition.Block))
            {
                var code = field.TypeAsString.StartsWith("TaxonomyFieldType", StringComparison.OrdinalIgnoreCase)
                    ? "TaxonomySchemaMappingUnavailable"
                    : field.TypeAsString.StartsWith("Lookup", StringComparison.OrdinalIgnoreCase)
                        ? "LookupSchemaMappingUnavailable"
                        : "SourceFieldSchemaUnavailable";
                issues.Add(Issue(code, $"field-schema:{field.InternalName}:{field.FieldId:D}", field.Reason));
            }

            var parentWillBeProvided = context?.WillProvideContentType(plan.ParentContentTypeId) == true;
            if ((!probe.ParentContentTypeAvailable
                    || !string.Equals(probe.ResolvedParentContentTypeId, plan.ParentContentTypeId, StringComparison.OrdinalIgnoreCase))
                && !parentWillBeProvided)
            {
                issues.Add(Issue("TargetContentTypeParentUnavailable",
                    $"target-parent-content-type:{plan.ParentContentTypeId}",
                    $"The target does not expose exact parent content type '{plan.ParentContentTypeName}' ({plan.ParentContentTypeId})."));
            }
            else if (parentWillBeProvided && !probe.ParentContentTypeAvailable)
            {
                warnings.Add($"The sealed content type closure creates parent '{plan.ParentContentTypeId}' before '{plan.ContentTypeId}'; execution still performs a strict fresh parent readback.");
            }

            foreach (var collisionId in probe.SameNameDifferentIds)
            {
                issues.Add(Issue("TargetContentTypeNameCollision", $"target-content-type-name:{plan.Name}",
                    $"The target already exposes content type name '{plan.Name}' under different ID '{collisionId}'."));
            }

            var targetFields = probe.Fields.ToDictionary(value => value.FieldId);
            var parentLinks = probe.ParentFieldLinks.ToDictionary(value => value.FieldId);
            foreach (var field in plan.Fields.Where(value => value.Disposition != FieldSchemaMaterializationDisposition.Block))
            {
                FieldSchemaTargetProbe target;
                targetFields.TryGetValue(field.FieldId, out target);
                if (field.TaxonomyMappingMode == TaxonomyTargetMappingMode.PreserveUnresolvedSourceReference)
                {
                    if (target?.UnresolvedTargetTermSetExists != false)
                    {
                        issues.Add(Issue(
                            target?.UnresolvedTargetTermSetExists == true
                                ? "TargetUnresolvedTaxonomyReferenceCollision"
                                : "TargetUnresolvedTaxonomyReferenceProbeUnavailable",
                            $"target-taxonomy-termset:{field.TargetTermSetId:D}",
                            target?.UnresolvedTargetTermSetExists == true
                                ? $"Target TermSet '{field.TargetTermSetId:D}' exists as '{target.UnresolvedTargetTermSetName}'; using it would heal the source-invalid taxonomy relationship."
                                : $"The selected unresolved target TermSet '{field.TargetTermSetId:D}' was not freshly probed."));
                    }
                    else
                    {
                        warnings.Add($"Taxonomy field '{field.InternalName}' ({field.FieldId:D}) intentionally targets absent TermSet '{field.TargetTermSetId:D}' while retaining source TermSet '{field.SourceTermSetId:D}' in the sealed plan; no TermSet asset will be created or repaired.");
                    }
                }
                if (field.Disposition == FieldSchemaMaterializationDisposition.RequireTargetRuntime)
                {
                    var generatedTaxonomyCompanion = IsGeneratedTaxonomyCompanion(plan, field);
                    var parentWillProvideField = field.Role == FieldSchemaRole.InheritedFromParent
                        && context?.WillProvideParentFieldLink(plan.ParentContentTypeId, field.FieldId) == true;
                    var featureWillProvideField = context?.WillProvisionRuntimeField(field.FieldId) == true;
                    if (target == null || !target.Exists)
                    {
                        if (parentWillProvideField)
                        {
                            warnings.Add($"The sealed parent content type transaction provides inherited field '{field.InternalName}' ({field.FieldId:D}) before creating '{plan.ContentTypeId}'.");
                        }
                        else if (featureWillProvideField)
                        {
                            warnings.Add($"A sealed platform-feature transaction provides target-runtime field '{field.InternalName}' ({field.FieldId:D}) before content type materialization.");
                        }
                        else
                        {
                            issues.Add(Issue("TargetFieldSchemaUnavailable", $"target-field:{field.FieldId:D}",
                                $"The target runtime does not expose required field '{field.InternalName}' ({field.FieldId:D})."));
                        }
                    }
                    else if ((!generatedTaxonomyCompanion
                            && !string.Equals(target.InternalName, field.InternalName, StringComparison.OrdinalIgnoreCase))
                        || !string.Equals(target.TypeAsString, field.TypeAsString, StringComparison.OrdinalIgnoreCase))
                    {
                        issues.Add(FieldCollision(field, target, "target-runtime field identity or type differs"));
                    }
                    else if (generatedTaxonomyCompanion
                        && !string.Equals(target.InternalName, field.InternalName, StringComparison.OrdinalIgnoreCase))
                    {
                        warnings.Add($"Target runtime generated taxonomy companion field '{target.InternalName}' for source companion '{field.InternalName}' ({field.FieldId:D}); GUID and Note type match, and the companion is not written directly.");
                    }
                    else if (field.Role == FieldSchemaRole.InheritedFromParent
                        && !parentLinks.ContainsKey(field.FieldId)
                        && !parentWillProvideField
                        && !featureWillProvideField)
                    {
                        issues.Add(Issue("TargetParentFieldLinkUnavailable",
                            $"target-parent-content-type-field-link:{field.FieldId:D}",
                            $"The target parent content type does not expose inherited field link '{field.InternalName}' ({field.FieldId:D})."));
                    }
                    else if (field.Role == FieldSchemaRole.InheritedFromParent
                        && !parentLinks.ContainsKey(field.FieldId)
                        && parentWillProvideField)
                    {
                        warnings.Add($"The sealed parent content type transaction supplies inherited field link '{field.InternalName}' ({field.FieldId:D}) before creating '{plan.ContentTypeId}'.");
                    }
                    else if (field.Role == FieldSchemaRole.InheritedFromParent
                        && !parentLinks.ContainsKey(field.FieldId)
                        && featureWillProvideField)
                    {
                        warnings.Add($"A sealed platform-feature transaction supplies inherited field link '{field.InternalName}' ({field.FieldId:D}) before creating '{plan.ContentTypeId}'.");
                    }

                    continue;
                }

                if (target != null && target.Exists
                    && (string.IsNullOrWhiteSpace(field.TargetPortableSchemaSha256)
                        || !string.Equals(target.PortableSchemaSha256, field.TargetPortableSchemaSha256, StringComparison.OrdinalIgnoreCase)))
                {
                    issues.Add(FieldCollision(field, target, "the same field GUID has different portable schema"));
                }
            }

            var reconcileExistingOwnedContentType = false;
            if (probe.ContentTypeExists)
            {
                var identityMetadataMatches = string.Equals(probe.ExistingName, plan.Name, StringComparison.Ordinal)
                    && string.Equals(probe.ExistingGroup ?? string.Empty, plan.Group ?? string.Empty, StringComparison.Ordinal)
                    && probe.ExistingReadOnly == plan.ReadOnly
                    && probe.ExistingSealed == plan.Sealed
                    && probe.ExistingHidden == plan.Hidden
                    && string.Equals(probe.ExistingParentContentTypeId, plan.ParentContentTypeId, StringComparison.OrdinalIgnoreCase);
                if (!identityMetadataMatches)
                {
                    issues.Add(Issue("TargetContentTypeCollision", $"target-content-type:{plan.ContentTypeId}",
                        "The exact content type ID exists with different metadata or parent lineage."));
                }
                var descriptionMatches = string.Equals(
                    probe.ExistingDescription ?? string.Empty,
                    plan.Description ?? string.Empty,
                    StringComparison.Ordinal);
                if (!descriptionMatches)
                {
                    if (identityMetadataMatches
                        && plan.Disposition == ContentTypeMaterializationDisposition.CreateOwned)
                    {
                        reconcileExistingOwnedContentType = true;
                        warnings.Add("The exact-ID interrupted create temporarily inherited its parent description; reconciliation will restore the reviewed description.");
                    }
                    else
                    {
                        issues.Add(Issue("TargetContentTypeCollision", $"target-content-type:{plan.ContentTypeId}",
                            "The exact content type ID exists with a different description."));
                    }
                }

                var existingLinks = probe.ExistingFieldLinks.ToDictionary(value => value.FieldId);
                foreach (var expected in plan.RequiredFieldLinks)
                {
                    ContentTypeFieldLinkTargetProbe actual;
                    if (!existingLinks.TryGetValue(expected.FieldId, out actual)
                        || actual.Required != expected.Required
                        || actual.Hidden != expected.Hidden)
                    {
                        if (identityMetadataMatches
                            && plan.Disposition == ContentTypeMaterializationDisposition.CreateOwned)
                        {
                            reconcileExistingOwnedContentType = true;
                            warnings.Add($"The exact-ID content type is an interrupted create and will reconcile field link '{expected.Name}' ({expected.FieldId:D}).");
                        }
                        else
                        {
                            issues.Add(Issue("TargetContentTypeFieldLinkCollision",
                                $"target-content-type-field-link:{expected.FieldId:D}",
                                $"The existing target content type does not expose exact required field link '{expected.Name}' ({expected.FieldId:D})."));
                        }
                    }
                }
            }
            else if (plan.Disposition == ContentTypeMaterializationDisposition.ReuseOwned)
            {
                issues.Add(Issue("TargetRuntimeContentTypeUnavailable", $"target-content-type:{plan.ContentTypeId}",
                    $"Partial source schema evidence permits exact target-runtime reuse only, but content type '{plan.Name}' ({plan.ContentTypeId}) does not exist at the target."));
            }
            else if (!probe.CanManageContentTypes)
            {
                issues.Add(Issue("TargetContentTypeWriteUnavailable", "target-content-type-schema-permission",
                    "The effective target permissions do not include ManageLists required for site-field and content-type creation."));
            }

            if (issues.Count == 0 && probe.ContentTypeExists)
            {
                warnings.Add(plan.Disposition == ContentTypeMaterializationDisposition.ReuseOwned
                    ? "Partial source schema evidence was admitted by exact target-runtime content type reuse; no field or content type schema will be created or repaired."
                    : "The exact required-field content type closure already exists and can be reused without a write.");
            }

            var disposition = probe.ContentTypeExists
                ? reconcileExistingOwnedContentType && issues.Count == 0
                    ? ContentTypeMaterializationDisposition.CreateOwned
                    : ContentTypeMaterializationDisposition.ReuseOwned
                : plan.Disposition == ContentTypeMaterializationDisposition.ReuseOwned
                    ? ContentTypeMaterializationDisposition.ReuseOwned
                    : ContentTypeMaterializationDisposition.CreateOwned;
            return Result(issues, warnings, disposition);
        }

        private static MigrationIssue FieldCollision(
            FieldSchemaMaterializationPlan field,
            FieldSchemaTargetProbe target,
            string reason)
        {
            return Issue("TargetFieldSchemaCollision", $"target-field-schema:{field.FieldId:D}",
                $"Target field collision for '{field.InternalName}' ({field.FieldId:D}): {reason}; target internalName='{target.InternalName}', type='{target.TypeAsString}'.");
        }

        private static bool IsGeneratedTaxonomyCompanion(
            ContentTypeMaterializationPlan plan,
            FieldSchemaMaterializationPlan field)
        {
            return field.Role == FieldSchemaRole.Dependency
                && field.Hidden
                && string.Equals(field.TypeAsString, "Note", StringComparison.OrdinalIgnoreCase)
                && plan.Fields.Any(value => value.HiddenTextFieldId == field.FieldId);
        }

        private static MigrationIssue Issue(string code, string ingredient, string message)
        {
            return new MigrationIssue
            {
                Code = code,
                Severity = MigrationIssueSeverity.Blocker,
                Subject = "ContentTypeSchema",
                Ingredient = ingredient,
                Message = message
            };
        }

        private static ContentTypeTargetAdmission Result(
            IEnumerable<MigrationIssue> issues,
            IEnumerable<string> warnings,
            ContentTypeMaterializationDisposition disposition)
        {
            var orderedIssues = issues
                .GroupBy(value => value.Code + "\u001f" + value.Ingredient + "\u001f" + value.Message, StringComparer.Ordinal)
                .Select(value => value.First())
                .OrderBy(value => value.Code, StringComparer.Ordinal)
                .ThenBy(value => value.Ingredient, StringComparer.Ordinal)
                .ThenBy(value => value.Message, StringComparer.Ordinal)
                .ToList();
            return new ContentTypeTargetAdmission
            {
                IsEligible = orderedIssues.Count == 0,
                Disposition = orderedIssues.Count == 0 ? disposition : ContentTypeMaterializationDisposition.Block,
                Issues = orderedIssues,
                Warnings = warnings.Distinct(StringComparer.Ordinal).OrderBy(value => value, StringComparer.Ordinal).ToList()
            };
        }
    }
}
