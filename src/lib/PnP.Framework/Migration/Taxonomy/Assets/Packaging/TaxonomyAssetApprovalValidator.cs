using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace PnP.Framework.Migration.Taxonomy.Assets.Packaging
{
    public static class TaxonomyAssetApprovalValidator
    {
        public static void Validate(
            TaxonomyAssetReviewPlan reviewPlan,
            TaxonomyAssetApprovalManifest manifest,
            bool requireDigest = true,
            bool requireAllDecided = true)
        {
            TaxonomyAssetReviewPlanValidator.Validate(reviewPlan, true, true);
            if (manifest == null)
            {
                throw new ArgumentNullException(nameof(manifest));
            }

            var errors = new List<string>();
            if (!string.Equals(manifest.SchemaVersion, "pnp-taxonomy-asset-approval/v1", StringComparison.Ordinal))
            {
                errors.Add("Unsupported taxonomy asset approval schema.");
            }
            if (!string.Equals(manifest.ReviewPlanDigest, reviewPlan.PlanDigest, StringComparison.OrdinalIgnoreCase))
            {
                errors.Add("The approval does not bind the reviewed taxonomy plan digest.");
            }
            if (manifest.GeneratedAtUtc == default(DateTimeOffset))
            {
                errors.Add("The approval template generation time is missing.");
            }
            if (requireDigest && (!manifest.ApprovedAtUtc.HasValue || string.IsNullOrWhiteSpace(manifest.ApprovedBy)))
            {
                errors.Add("The sealed approval identity or timestamp is missing.");
            }

            var expected = ExpectedActions(reviewPlan);
            var actual = new Dictionary<string, TaxonomyAssetActionApproval>(StringComparer.Ordinal);
            foreach (var action in (manifest.Actions ?? new List<TaxonomyAssetActionApproval>()).Where(value => value != null))
            {
                if (string.IsNullOrWhiteSpace(action.ActionId) || actual.ContainsKey(action.ActionId))
                {
                    errors.Add("The approval contains a blank or duplicate action ID.");
                    continue;
                }
                actual[action.ActionId] = action;
                if (!expected.TryGetValue(action.ActionId, out var reviewed))
                {
                    errors.Add("The approval contains an action absent from the reviewed plan: '" + action.ActionId + "'.");
                    continue;
                }
                ValidateAction(reviewed, action, requireAllDecided, errors);
            }
            foreach (var missing in expected.Keys.Except(actual.Keys, StringComparer.Ordinal))
            {
                errors.Add("The approval omits reviewed action '" + missing + "'.");
            }

            ValidateDependencies(reviewPlan, actual, errors);
            if (requireDigest && (string.IsNullOrWhiteSpace(manifest.ApprovalDigest)
                || !string.Equals(
                    manifest.ApprovalDigest,
                    TaxonomyAssetApprovalFactory.ComputeDigest(manifest),
                    StringComparison.OrdinalIgnoreCase)))
            {
                errors.Add("The taxonomy asset approval digest is absent or invalid.");
            }
            if (errors.Count > 0)
            {
                throw new InvalidDataException("Invalid taxonomy asset approval: " + string.Join(" ", errors));
            }
        }

        private static IDictionary<string, TaxonomyAssetActionApproval> ExpectedActions(TaxonomyAssetReviewPlan reviewPlan)
        {
            var result = new Dictionary<string, TaxonomyAssetActionApproval>(StringComparer.Ordinal);
            var groupPlans = reviewPlan.TermGroups.ToDictionary(
                value => TaxonomyAssetApprovalFactory.GroupKey(value.Source.TenantId, value.Source.TermStoreId),
                StringComparer.Ordinal);
            foreach (var probe in reviewPlan.TermGroupProbes)
            {
                var plan = groupPlans[TaxonomyAssetApprovalFactory.GroupKey(probe.SourceTenantId, probe.SourceTermStoreId)];
                var id = TaxonomyAssetApprovalFactory.TermGroupActionId(probe.SourceTenantId, probe.SourceTermStoreId);
                result[id] = new TaxonomyAssetActionApproval
                {
                    ActionId = id,
                    Kind = TaxonomyAssetKind.TermGroup,
                    SourceTenantId = probe.SourceTenantId,
                    SourceTermStoreId = probe.SourceTermStoreId,
                    TargetTermStoreId = probe.TargetTermStoreId,
                    TargetTermGroupId = probe.ResolvedTargetGroupId ?? plan.PreferredTargetGroupId,
                    ReviewedDisposition = probe.Disposition,
                    RequiresExplicitReview = probe.Disposition != TaxonomyAssetTargetDisposition.ReuseOwned
                };
            }
            var setPlans = reviewPlan.TermSets.ToDictionary(
                value => TaxonomyAssetApprovalFactory.SetKey(value.Source.TermStoreId, value.Source.TermSetId),
                StringComparer.Ordinal);
            foreach (var probe in reviewPlan.TermSetProbes)
            {
                var plan = setPlans[TaxonomyAssetApprovalFactory.SetKey(probe.SourceTermStoreId, probe.SourceTermSetId)];
                var id = TaxonomyAssetApprovalFactory.TermSetActionId(probe.SourceTermStoreId, probe.SourceTermSetId);
                result[id] = new TaxonomyAssetActionApproval
                {
                    ActionId = id,
                    Kind = TaxonomyAssetKind.TermSet,
                    SourceTenantId = plan.Source.TenantId,
                    SourceTermStoreId = probe.SourceTermStoreId,
                    SourceTermSetId = probe.SourceTermSetId,
                    TargetTermStoreId = probe.TargetTermStoreId,
                    TargetTermGroupId = plan.TargetGroupId,
                    TargetTermSetId = probe.ResolvedTargetTermSetId ?? plan.PreferredTargetTermSetId,
                    ReviewedDisposition = probe.Disposition,
                    RequiresExplicitReview = probe.Disposition != TaxonomyAssetTargetDisposition.ReuseOwned
                };
            }
            var termPlans = reviewPlan.Terms.ToDictionary(
                value => TaxonomyAssetApprovalFactory.TermKey(
                    value.Source.TermStoreId,
                    value.Source.TermSetId,
                    value.Source.TermId),
                StringComparer.Ordinal);
            foreach (var probe in reviewPlan.TermProbes)
            {
                var plan = termPlans[TaxonomyAssetApprovalFactory.TermKey(
                    probe.SourceTermStoreId,
                    probe.SourceTermSetId,
                    probe.SourceTermId)];
                var id = TaxonomyAssetApprovalFactory.TermActionId(
                    probe.SourceTermStoreId,
                    probe.SourceTermSetId,
                    probe.SourceTermId);
                result[id] = new TaxonomyAssetActionApproval
                {
                    ActionId = id,
                    Kind = TaxonomyAssetKind.Term,
                    SourceTenantId = plan.Source.TenantId,
                    SourceTermStoreId = probe.SourceTermStoreId,
                    SourceTermSetId = probe.SourceTermSetId,
                    SourceTermId = probe.SourceTermId,
                    TargetTermStoreId = probe.TargetTermStoreId,
                    TargetTermSetId = probe.TargetTermSetId,
                    TargetTermId = probe.ResolvedTargetTermId ?? plan.PreferredTargetTermId,
                    ReviewedDisposition = probe.Disposition,
                    RequiresExplicitReview = probe.Disposition != TaxonomyAssetTargetDisposition.ReuseOwned
                };
            }
            return result;
        }

        private static void ValidateAction(
            TaxonomyAssetActionApproval expected,
            TaxonomyAssetActionApproval actual,
            bool requireAllDecided,
            ICollection<string> errors)
        {
            if (actual.Kind != expected.Kind
                || actual.SourceTenantId != expected.SourceTenantId
                || actual.SourceTermStoreId != expected.SourceTermStoreId
                || actual.SourceTermSetId != expected.SourceTermSetId
                || actual.SourceTermId != expected.SourceTermId
                || actual.TargetTermStoreId != expected.TargetTermStoreId
                || actual.TargetTermGroupId != expected.TargetTermGroupId
                || actual.TargetTermSetId != expected.TargetTermSetId
                || actual.TargetTermId != expected.TargetTermId
                || actual.ReviewedDisposition != expected.ReviewedDisposition
                || actual.RequiresExplicitReview != expected.RequiresExplicitReview)
            {
                errors.Add("Approval action '" + actual.ActionId + "' differs from the reviewed target identity or disposition.");
            }
            if (!Enum.IsDefined(typeof(TaxonomyAssetKind), actual.Kind)
                || !Enum.IsDefined(typeof(TaxonomyAssetTargetDisposition), actual.ReviewedDisposition)
                || !Enum.IsDefined(typeof(TaxonomyAssetApprovalDecision), actual.Decision))
            {
                errors.Add("Approval action '" + actual.ActionId + "' contains an undefined kind, disposition, or decision.");
            }
            if (requireAllDecided && actual.Decision == TaxonomyAssetApprovalDecision.Pending)
            {
                errors.Add("Approval action '" + actual.ActionId + "' is still pending.");
            }
            if (actual.Decision == TaxonomyAssetApprovalDecision.Approve
                && !IsApprovable(actual.ReviewedDisposition))
            {
                errors.Add("Approval action '" + actual.ActionId + "' cannot approve disposition " + actual.ReviewedDisposition + ".");
            }
            if (actual.ExternalMutationApproved
                && (actual.Kind != TaxonomyAssetKind.Term
                    || actual.ReviewedDisposition != TaxonomyAssetTargetDisposition.CreateMissingAfterExternalApproval
                    || actual.Decision != TaxonomyAssetApprovalDecision.Approve))
            {
                errors.Add("External mutation is authorized only for an approved CreateMissingAfterExternalApproval Term action.");
            }
            if (actual.Decision == TaxonomyAssetApprovalDecision.Approve
                && actual.ReviewedDisposition == TaxonomyAssetTargetDisposition.CreateMissingAfterExternalApproval
                && !actual.ExternalMutationApproved)
            {
                errors.Add("Approval action '" + actual.ActionId + "' requires explicit external mutation authorization.");
            }
        }

        private static void ValidateDependencies(
            TaxonomyAssetReviewPlan reviewPlan,
            IDictionary<string, TaxonomyAssetActionApproval> actions,
            ICollection<string> errors)
        {
            var groups = reviewPlan.TermGroups.ToDictionary(
                value => TaxonomyAssetApprovalFactory.GroupKey(value.Source.TenantId, value.Source.TermStoreId),
                StringComparer.Ordinal);
            foreach (var set in reviewPlan.TermSets)
            {
                var setActionId = TaxonomyAssetApprovalFactory.TermSetActionId(
                    set.Source.TermStoreId,
                    set.Source.TermSetId);
                if (!actions.TryGetValue(setActionId, out var setAction)
                    || setAction.Decision != TaxonomyAssetApprovalDecision.Approve
                    || setAction.ReviewedDisposition == TaxonomyAssetTargetDisposition.ReviewExternalReuse)
                {
                    continue;
                }
                var groupKey = TaxonomyAssetApprovalFactory.GroupKey(set.Source.TenantId, set.Source.TermStoreId);
                if (!groups.TryGetValue(groupKey, out var group))
                {
                    errors.Add("Approved TermSet action '" + setActionId + "' has no deterministic TermGroup plan.");
                    continue;
                }
                var groupActionId = TaxonomyAssetApprovalFactory.TermGroupActionId(
                    group.Source.TenantId,
                    group.Source.TermStoreId);
                if (!actions.TryGetValue(groupActionId, out var groupAction)
                    || groupAction.Decision != TaxonomyAssetApprovalDecision.Approve)
                {
                    errors.Add("Approved TermSet action '" + setActionId + "' requires its TermGroup action to be approved.");
                }
            }

            var terms = reviewPlan.Terms.ToDictionary(
                value => TaxonomyAssetApprovalFactory.TermKey(
                    value.Source.TermStoreId,
                    value.Source.TermSetId,
                    value.Source.TermId),
                StringComparer.Ordinal);
            foreach (var term in terms.Values)
            {
                var actionId = TaxonomyAssetApprovalFactory.TermActionId(
                    term.Source.TermStoreId,
                    term.Source.TermSetId,
                    term.Source.TermId);
                if (!actions.TryGetValue(actionId, out var action)
                    || action.Decision != TaxonomyAssetApprovalDecision.Approve)
                {
                    continue;
                }
                var setActionId = TaxonomyAssetApprovalFactory.TermSetActionId(
                    term.Source.TermStoreId,
                    term.Source.TermSetId);
                if (!actions.TryGetValue(setActionId, out var setAction)
                    || setAction.Decision != TaxonomyAssetApprovalDecision.Approve)
                {
                    errors.Add("Approved Term action '" + actionId + "' requires its TermSet action to be approved.");
                }
                if (term.TargetParentTermId.HasValue)
                {
                    var parentActionId = TaxonomyAssetApprovalFactory.TermActionId(
                        term.Source.TermStoreId,
                        term.Source.TermSetId,
                        term.TargetParentTermId.Value);
                    if (!actions.TryGetValue(parentActionId, out var parentAction)
                        || parentAction.Decision != TaxonomyAssetApprovalDecision.Approve)
                    {
                        errors.Add("Approved Term action '" + actionId + "' requires its parent Term action to be approved.");
                    }
                }
                if (action.ReviewedDisposition == TaxonomyAssetTargetDisposition.CreateMissingAfterExternalApproval
                    && (!actions.TryGetValue(setActionId, out var externalSet)
                        || externalSet.ReviewedDisposition != TaxonomyAssetTargetDisposition.ReviewExternalReuse))
                {
                    errors.Add("External child creation requires an approved external TermSet reuse action.");
                }
            }
        }

        private static bool IsApprovable(TaxonomyAssetTargetDisposition disposition)
        {
            return disposition == TaxonomyAssetTargetDisposition.CreateMissing
                || disposition == TaxonomyAssetTargetDisposition.ReuseOwned
                || disposition == TaxonomyAssetTargetDisposition.ReconcileOwnedPlanDrift
                || disposition == TaxonomyAssetTargetDisposition.ReviewExternalReuse
                || disposition == TaxonomyAssetTargetDisposition.CreateMissingAfterExternalApproval;
        }
    }
}
