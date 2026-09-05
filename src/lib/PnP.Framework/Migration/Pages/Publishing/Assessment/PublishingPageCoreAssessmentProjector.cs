using PnP.Framework.Migration.Pages.Assessment;
using PnP.Framework.Migration.Pages.Capture;
using PnP.Framework.Migration.Pages.Fields;
using PnP.Framework.Migration.Pages.Fields.Taxonomy;
using PnP.Framework.Migration.Pages.Ingredients;
using PnP.Framework.Migration.Pages.Planning;
using PnP.Framework.Migration.Pages.Publishing.Ingredients;
using PnP.Framework.Migration.Pages.Publishing.Lifecycle;
using PnP.Framework.Migration.Pages.Runtime;
using PnP.Framework.Migration.Schema.Fields;
using PnP.Framework.Migration.Taxonomy;
using PnP.Framework.Migration.Taxonomy.Assets;
using System;
using System.Collections.Generic;
using System.Linq;

namespace PnP.Framework.Migration.Pages.Publishing.Assessment
{
    internal static class PublishingPageCoreAssessmentProjector
    {
        public static void Project(
            PublishingPageAssessmentContext context,
            PublishingPageAssessmentAccumulator assessments)
        {
            AddRuntimeAndArtifact(context, assessments);
            AddContentSecurityAndLifecycle(context, assessments);
            AddFields(context, assessments);
        }

        private static void AddRuntimeAndArtifact(
            PublishingPageAssessmentContext context,
            PublishingPageAssessmentAccumulator assessments)
        {
            var snapshot = context.Snapshot;
            var publishingRuntime = string.Equals(
                snapshot.Runtime?.AdapterId,
                PageRuntimeAdapterIds.Publishing,
                StringComparison.Ordinal);
            assessments.Add(
                PublishingPageIngredientIds.Runtime,
                publishingRuntime
                    ? PageIngredientAssessmentState.Determined
                    : PageIngredientAssessmentState.KnownGap,
                publishingRuntime ? IngredientCapability.Available : IngredientCapability.Incompatible,
                publishingRuntime ? IngredientDisposition.Preserve : IngredientDisposition.Defer,
                publishingRuntime ? "reuse-target-runtime" : "none",
                "policy.runtime.publishing",
                publishingRuntime
                    ? "The Publishing CLR runtime selects the Publishing adapter before target inspection."
                    : $"Runtime adapter '{snapshot.Runtime?.AdapterId ?? PageRuntimeAdapterIds.Unknown}' is not executable by the Publishing importer.",
                publishingRuntime ? PageRuntimeAdapterIds.Publishing : null,
                publishingRuntime ? null : "PublishingRuntimeUnavailable",
                "The target page resolves through the Publishing runtime without an error shell.");

            assessments.Add(
                PublishingPageIngredientIds.PageArtifact,
                publishingRuntime
                    ? PageIngredientAssessmentState.TargetInspectionRequired
                    : PageIngredientAssessmentState.KnownGap,
                publishingRuntime ? IngredientCapability.Available : IngredientCapability.Incompatible,
                publishingRuntime ? IngredientDisposition.Transform : IngredientDisposition.Defer,
                publishingRuntime ? "create-target-page-shell" : "none",
                "policy.page-artifact.publishing",
                publishingRuntime
                    ? "Create the target Publishing page shell at the exact mapped relative path; retain source ASPX bytes as immutable evidence."
                    : "The target page shell cannot be selected until a compatible page runtime is available.",
                publishingRuntime ? context.TargetPageServerRelativeUrl : null,
                publishingRuntime ? null : "PublishingRuntimeUnavailable",
                "The exact target page path exists after mutation.",
                "The source ASPX artifact remains digest-verifiable in the package.");
        }

        private static void AddContentSecurityAndLifecycle(
            PublishingPageAssessmentContext context,
            PublishingPageAssessmentAccumulator assessments)
        {
            var snapshot = context.Snapshot;
            var transformed = context.Replacements.Count > 0;
            assessments.Add(
                PublishingPageIngredientIds.PublishingContent,
                PageIngredientAssessmentState.Determined,
                IngredientCapability.Available,
                transformed ? IngredientDisposition.Transform : IngredientDisposition.Preserve,
                transformed ? "copy-and-rewrite-approved-references" : "copy-exact-value",
                "policy.content.publishing",
                transformed
                    ? "Replay PublishingPageContent after applying only the deterministic source-to-target reference rewrites."
                    : "Replay the captured PublishingPageContent without a text rewrite.",
                Anchor(context.TargetPageServerRelativeUrl, "PublishingPageContent"),
                null,
                "Fresh readback verifies the normalized PublishingPageContent digest.");

            var uniqueSecurity = snapshot.Security?.HasUniqueRoleAssignments == true;
            var securityBlocked = uniqueSecurity && context.Options.RequireInheritedPermissions;
            assessments.Add(
                PublishingPageIngredientIds.Security,
                securityBlocked
                    ? PageIngredientAssessmentState.KnownGap
                    : PageIngredientAssessmentState.Determined,
                securityBlocked ? IngredientCapability.Incompatible : IngredientCapability.Available,
                securityBlocked
                    ? IngredientDisposition.Defer
                    : uniqueSecurity ? IngredientDisposition.Delegate : IngredientDisposition.Preserve,
                securityBlocked
                    ? "none"
                    : uniqueSecurity ? "retain-snapshot" : "reuse-target-inheritance",
                "policy.security.page",
                securityBlocked
                    ? "The source page has unique permissions, but the selected policy requires inherited target permissions."
                    : uniqueSecurity
                        ? "Unique permissions remain fully captured for a separately reviewed principal-mapping transaction."
                        : "Reuse target permission inheritance.",
                securityBlocked ? null : Anchor(context.TargetPageServerRelativeUrl, "security"),
                securityBlocked ? "UniquePagePermissionsMappingRequired" : null,
                uniqueSecurity
                    ? "The exact captured role-assignment evidence remains in the snapshot."
                    : "The target page inherits permissions.");

            var sourcePublished = string.Equals(snapshot.Lifecycle?.Level, "Published", StringComparison.OrdinalIgnoreCase);
            var lifecyclePreserved = sourcePublished == (context.TargetLifecycle == PublishingPageTargetLifecycle.Published);
            assessments.Add(
                PublishingPageIngredientIds.Lifecycle,
                PageIngredientAssessmentState.Determined,
                IngredientCapability.Available,
                lifecyclePreserved ? IngredientDisposition.Preserve : IngredientDisposition.Transform,
                context.TargetLifecycle == PublishingPageTargetLifecycle.Published
                    ? "publish-after-verification"
                    : "check-in-as-draft",
                "policy.lifecycle.publishing",
                context.LifecycleReason,
                Anchor(context.TargetPageServerRelativeUrl, "lifecycle"),
                null,
                $"The final target lifecycle is '{context.TargetLifecycle}'.");
        }

        private static void AddFields(
            PublishingPageAssessmentContext context,
            PublishingPageAssessmentAccumulator assessments)
        {
            foreach (var field in context.Snapshot.Fields
                         .Where(value => value != null)
                         .OrderBy(value => value.InternalName, StringComparer.Ordinal))
            {
                var ingredientId = PublishingPageIngredientIds.Field(field.InternalName);
                if (context.WorkflowPolicy.FieldsHandledByPageWriter.Contains(field.InternalName))
                {
                    assessments.Add(
                        ingredientId,
                        PageIngredientAssessmentState.Determined,
                        IngredientCapability.Available,
                        IngredientDisposition.Preserve,
                        "handled-by-page-writer",
                        "policy.field.already-handled",
                        "The page writer owns this field explicitly.",
                        field.InternalName,
                        null,
                        $"The page writer verifies field '{field.InternalName}'.");
                    continue;
                }

                if (field.CaptureStatus is PageCaptureStatus.Failed or PageCaptureStatus.NotReturned)
                {
                    assessments.Add(
                        ingredientId,
                        PageIngredientAssessmentState.KnownGap,
                        IngredientCapability.Missing,
                        IngredientDisposition.Defer,
                        "none",
                        "policy.field.capture-evidence",
                        "The field definition remains captured, but no restorable current value was returned.",
                        mitigationCode: "PageFieldCaptureUnavailable");
                    AddTaxonomyRelationshipsAsMitigation(field, assessments, "PageFieldCaptureUnavailable",
                        "The owning taxonomy field has no complete restorable capture evidence.");
                    continue;
                }

                if (!field.HasValue
                    || field.Kind == PageFieldValueKind.Null
                    || (IsTaxonomy(field.Kind) && field.TaxonomyValues.Count == 0))
                {
                    assessments.Add(
                        ingredientId,
                        PageIngredientAssessmentState.Determined,
                        IngredientCapability.Available,
                        IngredientDisposition.Preserve,
                        "ensure-schema-without-value",
                        "policy.field.empty",
                        "The source item has no value for this field; preserve that absence and retain its schema evidence.",
                        field.InternalName,
                        null,
                        $"The target does not fabricate a value for field '{field.InternalName}'.");
                    continue;
                }

                var recognized = context.WorkflowPolicy.RecognizedPageFields.Contains(field.InternalName);
                if (!recognized)
                {
                    var targetRuntime = FieldOwnershipClassifier.IsTargetRuntime(field.Id, field.SchemaXml);
                    assessments.Add(
                        ingredientId,
                        targetRuntime
                            ? PageIngredientAssessmentState.TargetInspectionRequired
                            : PageIngredientAssessmentState.Determined,
                        targetRuntime ? IngredientCapability.Available : IngredientCapability.Unknown,
                        targetRuntime ? IngredientDisposition.Substitute : IngredientDisposition.Delegate,
                        targetRuntime ? "reuse-target-runtime-value" : "retain-snapshot",
                        targetRuntime ? "policy.field.target-runtime" : "policy.field.evidence-only",
                        targetRuntime
                            ? "The field is SharePoint-owned; target inspection must prove an equivalent runtime field before its source value is substituted."
                            : "The importer does not recognize this field yet. Its complete value remains in the snapshot and is not replayed.",
                        targetRuntime ? field.InternalName : null,
                        null,
                        targetRuntime
                            ? $"Fresh target inspection proves same-name, same-type runtime field '{field.InternalName}'."
                            : $"The reviewed evidence-only decision remains sealed for field '{field.InternalName}'.");
                    AddTaxonomyRelationshipsAsEvidence(field, assessments);
                    continue;
                }

                if (field.ReadOnly || string.Equals(field.TypeAsString, "Calculated", StringComparison.OrdinalIgnoreCase))
                {
                    assessments.Add(
                        ingredientId,
                        PageIngredientAssessmentState.Determined,
                        IngredientCapability.Available,
                        IngredientDisposition.Substitute,
                        "target-runtime-value",
                        "policy.field.target-runtime",
                        field.ReadOnly
                            ? "The recognized source field is read-only; SharePoint owns the target value."
                            : "The recognized calculated field is recomputed by SharePoint.",
                        field.InternalName,
                        null,
                        $"Target runtime owns field '{field.InternalName}' after import.");
                    AddTaxonomyRelationshipsAsEvidence(field, assessments);
                    continue;
                }

                if (IsTaxonomy(field.Kind))
                {
                    AssessTaxonomyField(context, field, assessments);
                    continue;
                }

                if (RequiresIdentityMapping(field.Kind))
                {
                    var required = field.Required;
                    assessments.Add(
                        ingredientId,
                        required
                            ? PageIngredientAssessmentState.KnownGap
                            : PageIngredientAssessmentState.Determined,
                        required ? IngredientCapability.Incompatible : IngredientCapability.Unknown,
                        required ? IngredientDisposition.Defer : IngredientDisposition.Delegate,
                        required ? "none" : "retain-snapshot",
                        "policy.field.identity-mapping",
                        required
                            ? "The required identity-bound value needs an explicit cross-site mapping."
                            : "No reviewed identity mapping exists; retain the optional value as evidence and leave the target unset.",
                        mitigationCode: required ? "PageFieldIdentityMappingUnavailable" : null);
                    continue;
                }

                if (!PageFieldPlanner.IsImportableKind(field.Kind))
                {
                    assessments.Add(
                        ingredientId,
                        PageIngredientAssessmentState.Determined,
                        IngredientCapability.Unknown,
                        IngredientDisposition.Delegate,
                        "retain-snapshot",
                        "policy.field.evidence-only",
                        $"No importer is registered for value kind '{field.Kind}'; retain the complete source value as evidence.",
                        verificationAssertions: $"The reviewed evidence-only decision remains sealed for field '{field.InternalName}'.");
                    continue;
                }

                assessments.Add(
                    ingredientId,
                    PageIngredientAssessmentState.TargetInspectionRequired,
                    IngredientCapability.Available,
                    IngredientDisposition.Preserve,
                    "map-one-to-one",
                    "policy.field.recognized",
                    "The recognized value is source-portable; target inspection must prove a writable same-name, same-type Pages field.",
                    field.InternalName,
                    null,
                    $"Fresh target inspection proves field '{field.InternalName}' is writable and type-compatible.",
                    $"Fresh readback round-trips the approved value for field '{field.InternalName}'.");
            }
        }

        private static void AssessTaxonomyField(
            PublishingPageAssessmentContext context,
            PageFieldValueSnapshot field,
            PublishingPageAssessmentAccumulator assessments)
        {
            var fieldErrors = PageTaxonomyRelationshipEvidence.ValidateSealedField(field).ToList();
            var binding = field.TaxonomyBinding;
            var boundCandidate = binding == null
                ? null
                : PagePlanningTaxonomyMappingResolver.FindAssessmentCandidate(
                    context.TaxonomyAssetReviewPlan,
                    binding.TermStoreId,
                    binding.BoundTermSetId);
            var boundMappings = binding == null
                ? Array.Empty<TaxonomyTargetMapping>()
                : context.Options.TaxonomySchemaMappings.Where(value =>
                    value.SourceTermStoreId == binding.TermStoreId
                    && value.SourceTermSetId == binding.BoundTermSetId).ToArray();
            var boundReviewRequired = boundMappings.Length == 0
                && boundCandidate?.Disposition == TaxonomyAssetTargetDisposition.ReviewExternalReuse;
            var mappingError = boundMappings.Length == 1
                ? null
                : boundMappings.Length == 0
                    ? boundReviewRequired
                        ? "The read-only taxonomy preflight found an external target Term Set candidate, but explicit reuse approval is still pending."
                        : "No deterministic target mapping exists for the source field's bound Term Store and Term Set."
                    : "More than one target mapping exists for the source field's bound Term Store and Term Set.";
            var relationshipMitigation = fieldErrors.Count > 0 || mappingError != null;
            var relationshipReviewRequired = boundReviewRequired;
            var reviewCandidates = new List<TaxonomyAssetMappingCandidate>();
            if (boundReviewRequired && boundCandidate != null)
            {
                reviewCandidates.Add(boundCandidate);
            }
            foreach (var value in field.TaxonomyValues.Where(item => item != null))
            {
                var errors = new List<string>(fieldErrors);
                var valueReviewRequired = boundReviewRequired;
                var reviewCandidate = boundReviewRequired ? boundCandidate : null;
                errors.AddRange(PageTaxonomyRelationshipEvidence.GetFidelityErrors(field, value));
                if (mappingError != null)
                {
                    errors.Add(mappingError);
                }
                if (value.Relationship?.State == TaxonomyRelationshipState.LiveOutsideBoundTermSet
                    && value.Relationship.LiveTermSetId.HasValue
                    && binding != null)
                {
                    var liveMappings = context.Options.TaxonomySchemaMappings.Count(candidate =>
                        candidate.SourceTermStoreId == binding.TermStoreId
                        && candidate.SourceTermSetId == value.Relationship.LiveTermSetId.Value);
                    if (liveMappings != 1)
                    {
                        var liveCandidate = PagePlanningTaxonomyMappingResolver.FindAssessmentCandidate(
                            context.TaxonomyAssetReviewPlan,
                            binding.TermStoreId,
                            value.Relationship.LiveTermSetId.Value);
                        if (liveMappings == 0
                            && liveCandidate?.Disposition == TaxonomyAssetTargetDisposition.ReviewExternalReuse)
                        {
                            errors.Add("The live outside-bound Term Set has one external target candidate, but explicit reuse approval is still pending.");
                            valueReviewRequired = true;
                            reviewCandidate = liveCandidate;
                            reviewCandidates.Add(liveCandidate);
                        }
                        else
                        {
                            errors.Add("The live outside-bound Term Set has no unique deterministic target mapping.");
                        }
                    }
                }

                Guid.TryParse(value.TermGuid, out var termId);
                var ingredientId = PublishingPageIngredientIds.TaxonomyRelationship(field.Id, termId, value.WssId);
                var uniqueErrors = errors.Where(error => !string.IsNullOrWhiteSpace(error))
                    .Distinct(StringComparer.Ordinal).ToArray();
                if (uniqueErrors.Length > 0)
                {
                    relationshipMitigation = true;
                    relationshipReviewRequired |= valueReviewRequired;
                    assessments.Add(
                        ingredientId,
                        PageIngredientAssessmentState.KnownGap,
                        IngredientCapability.Incompatible,
                        IngredientDisposition.Defer,
                        valueReviewRequired
                            ? "review-external-term-set-reuse-then-reproduce-captured-relationship"
                            : "collect-deterministic-taxonomy-relationship-evidence",
                        "policy.taxonomy-relationship.fidelity",
                        string.Join(" ", uniqueErrors),
                        TaxonomyReviewTargetIdentity(reviewCandidate, termId),
                        mitigationCode: valueReviewRequired
                            ? "TaxonomyExternalReuseApprovalRequired"
                            : mappingError != null
                                ? "TaxonomyMappingUnavailable"
                            : "TaxonomyRelationshipEvidenceUnavailable",
                        verificationAssertions: TaxonomyReviewAssertions(
                            reviewCandidate,
                            termId,
                            value.Relationship?.State));
                    continue;
                }

                var mapping = boundMappings[0];
                assessments.Add(
                    ingredientId,
                    PageIngredientAssessmentState.TargetInspectionRequired,
                    IngredientCapability.Available,
                    IngredientDisposition.Transform,
                    TaxonomyRealization(value.Relationship.State),
                    "policy.taxonomy-relationship.fidelity",
                    "The exact captured relationship state is eligible for reproduction without creating, substituting, or repairing a Term; target taxonomy evidence must still be inspected.",
                    mapping.TargetTermStoreId.ToString("D") + "/" + mapping.TargetTermSetId.ToString("D") + "/" + termId.ToString("D"),
                    null,
                    "Fresh target inspection proves that replay preserves the exact live, outside-bound, or dangling relationship state.",
                    "Fresh readback verifies the exact Term GUID and target-local hidden-list relationship.");
            }

            assessments.Add(
                PublishingPageIngredientIds.Field(field.InternalName),
                relationshipMitigation
                    ? PageIngredientAssessmentState.KnownGap
                    : PageIngredientAssessmentState.TargetInspectionRequired,
                relationshipMitigation ? IngredientCapability.Incompatible : IngredientCapability.Available,
                relationshipMitigation ? IngredientDisposition.Defer : IngredientDisposition.Transform,
                relationshipMitigation
                    ? relationshipReviewRequired
                        ? "review-external-taxonomy-dependencies-then-reproduce-captured-relationships"
                        : "collect-deterministic-taxonomy-relationship-evidence"
                    : "reproduce-reviewed-taxonomy-relationships",
                "policy.field.taxonomy-relationships",
                relationshipMitigation
                    ? relationshipReviewRequired
                        ? "One or more captured taxonomy relationship ingredients are waiting for explicit external target reuse approval."
                        : "One or more captured taxonomy relationship ingredients lack complete evidence or a unique deterministic mapping."
                    : "Every captured taxonomy relationship has a source-authoritative candidate action; target field and term-state inspection remains required.",
                relationshipMitigation
                    ? TaxonomyFieldReviewTargetIdentity(field.InternalName, reviewCandidates)
                    : field.InternalName,
                relationshipMitigation
                    ? relationshipReviewRequired
                        ? "TaxonomyRelationshipApprovalRequired"
                        : "TaxonomyRelationshipActionUnavailable"
                    : null,
                verificationAssertions: relationshipMitigation
                    ? TaxonomyFieldReviewAssertions(field.InternalName, reviewCandidates)
                    : new[]
                    {
                        $"Fresh target inspection proves taxonomy field '{field.InternalName}' has the reviewed binding and companion field."
                    });
        }

        private static string TaxonomyReviewTargetIdentity(
            TaxonomyAssetMappingCandidate candidate,
            Guid termId)
        {
            if (candidate == null)
            {
                return null;
            }

            return candidate.TargetTermStoreId.ToString("D") + "/"
                + candidate.TargetTermSetId.ToString("D") + "/"
                + termId.ToString("D");
        }

        private static string TaxonomyFieldReviewTargetIdentity(
            string fieldInternalName,
            IEnumerable<TaxonomyAssetMappingCandidate> candidates)
        {
            var targets = DistinctReviewCandidates(candidates)
                .Select(value => value.TargetTermStoreId.ToString("D") + "/" + value.TargetTermSetId.ToString("D"))
                .ToArray();
            return targets.Length == 0
                ? fieldInternalName
                : fieldInternalName + " -> " + string.Join(",", targets);
        }

        private static string[] TaxonomyReviewAssertions(
            TaxonomyAssetMappingCandidate candidate,
            Guid termId,
            TaxonomyRelationshipState? relationshipState)
        {
            if (candidate == null)
            {
                return new[]
                {
                    "A later evidence pass must resolve exactly one deterministic target taxonomy mapping before execution.",
                    "Fresh readback must preserve the captured relationship state and exact Term GUID without substitution or repair."
                };
            }

            return (candidate.VerificationAssertions ?? Array.Empty<string>())
                .Concat(new[]
                {
                    "Approval must bind this relationship to the exact external target TermSet before any target mutation.",
                    $"Fresh readback preserves captured relationship state '{relationshipState}' and exact Term GUID '{termId:D}' without substitution or repair."
                })
                .Where(value => !string.IsNullOrWhiteSpace(value))
                .Distinct(StringComparer.Ordinal)
                .ToArray();
        }

        private static string[] TaxonomyFieldReviewAssertions(
            string fieldInternalName,
            IEnumerable<TaxonomyAssetMappingCandidate> candidates)
        {
            var assertions = DistinctReviewCandidates(candidates)
                .SelectMany(value => value.VerificationAssertions ?? Array.Empty<string>())
                .Concat(new[]
                {
                    $"No taxonomy relationship for field '{fieldInternalName}' is executed until every external TermSet dependency has explicit digest-bound approval.",
                    $"Fresh readback preserves every captured live, outside-bound, missing, or dangling relationship for field '{fieldInternalName}' without substituting or repairing a Term."
                })
                .Where(value => !string.IsNullOrWhiteSpace(value))
                .Distinct(StringComparer.Ordinal)
                .ToArray();
            return assertions;
        }

        private static IEnumerable<TaxonomyAssetMappingCandidate> DistinctReviewCandidates(
            IEnumerable<TaxonomyAssetMappingCandidate> candidates)
        {
            return (candidates ?? Enumerable.Empty<TaxonomyAssetMappingCandidate>())
                .Where(value => value != null)
                .GroupBy(
                    value => value.TargetTermStoreId.ToString("D") + "/" + value.TargetTermSetId.ToString("D"),
                    StringComparer.Ordinal)
                .Select(group => group.First())
                .OrderBy(value => value.TargetTermStoreId)
                .ThenBy(value => value.TargetTermSetId);
        }

        private static void AddTaxonomyRelationshipsAsEvidence(
            PageFieldValueSnapshot field,
            PublishingPageAssessmentAccumulator assessments)
        {
            foreach (var value in field.TaxonomyValues.Where(item => item != null))
            {
                Guid.TryParse(value.TermGuid, out var termId);
                assessments.Add(
                    PublishingPageIngredientIds.TaxonomyRelationship(field.Id, termId, value.WssId),
                    PageIngredientAssessmentState.Determined,
                    IngredientCapability.Unknown,
                    IngredientDisposition.Delegate,
                    "retain-sealed-relationship-evidence",
                    "policy.taxonomy-relationship.evidence-only",
                    "The owning field is not selected for replay; preserve the exact relationship snapshot without repair.",
                    verificationAssertions: "The taxonomy relationship evidence remains digest-verifiable in the source snapshot.");
            }
        }

        private static void AddTaxonomyRelationshipsAsMitigation(
            PageFieldValueSnapshot field,
            PublishingPageAssessmentAccumulator assessments,
            string code,
            string reason)
        {
            foreach (var value in field.TaxonomyValues.Where(item => item != null))
            {
                Guid.TryParse(value.TermGuid, out var termId);
                assessments.Add(
                    PublishingPageIngredientIds.TaxonomyRelationship(field.Id, termId, value.WssId),
                    PageIngredientAssessmentState.KnownGap,
                    IngredientCapability.Missing,
                    IngredientDisposition.Defer,
                    "none",
                    "policy.taxonomy-relationship.capture-evidence",
                    reason,
                    mitigationCode: code);
            }
        }

        private static bool IsTaxonomy(PageFieldValueKind kind)
        {
            return kind is PageFieldValueKind.Taxonomy or PageFieldValueKind.TaxonomyCollection;
        }

        private static bool RequiresIdentityMapping(PageFieldValueKind kind)
        {
            return kind is PageFieldValueKind.User
                or PageFieldValueKind.UserCollection
                or PageFieldValueKind.Lookup
                or PageFieldValueKind.LookupCollection;
        }

        private static string TaxonomyRealization(TaxonomyRelationshipState state)
        {
            return state == TaxonomyRelationshipState.LiveInBoundTermSet
                ? "reuse-live-term-in-mapped-bound-set"
                : state == TaxonomyRelationshipState.LiveOutsideBoundTermSet
                    ? "reproduce-live-term-outside-mapped-bound-set"
                    : state == TaxonomyRelationshipState.DanglingTermAbsent
                        ? "reproduce-dangling-term-with-target-local-wssid"
                        : "none";
        }

        private static string Anchor(string path, string fragment)
        {
            return string.IsNullOrWhiteSpace(path) ? null : path + "#" + fragment;
        }
    }
}
