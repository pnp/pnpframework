using PnP.Framework.Migration.Execution;
using PnP.Framework.Migration.Taxonomy.Assets.Packaging;
using System;
using System.Collections.Generic;
using System.Linq;

namespace PnP.Framework.Migration.Taxonomy.Assets.Execution
{
    public static class TaxonomyAssetExecutionAdmissionEvaluator
    {
        public static TaxonomyAssetExecutionAdmission Evaluate(
            TaxonomyAssetReviewPlan reviewedPlan,
            TaxonomyAssetReviewPlan freshInspection,
            TaxonomyAssetApprovalManifest approval)
        {
            TaxonomyAssetReviewPlanValidator.Validate(reviewedPlan, true, true);
            TaxonomyAssetReviewPlanValidator.Validate(freshInspection, true, true);
            TaxonomyAssetApprovalValidator.Validate(reviewedPlan, approval, true, true);

            var result = new TaxonomyAssetExecutionAdmission
            {
                ReviewPlanDigest = reviewedPlan.PlanDigest,
                ApprovalDigest = approval.ApprovalDigest,
                FreshInspectionPlanDigest = freshInspection.PlanDigest
            };
            if (!string.Equals(reviewedPlan.SourceSnapshotDigest, freshInspection.SourceSnapshotDigest, StringComparison.OrdinalIgnoreCase)
                || reviewedPlan.TargetTermStoreId != freshInspection.TargetTermStoreId)
            {
                AddFailure(result, "TaxonomyInspectionBoundaryChanged", "taxonomy", "Fresh inspection changed the source snapshot or target TermStore boundary.");
            }

            var approvals = approval.Actions.ToDictionary(value => value.ActionId, StringComparer.Ordinal);
            var reviewedGroupProbes = reviewedPlan.TermGroupProbes.ToDictionary(
                value => TaxonomyAssetApprovalFactory.GroupKey(value.SourceTenantId, value.SourceTermStoreId),
                StringComparer.Ordinal);
            var freshGroupProbes = freshInspection.TermGroupProbes.ToDictionary(
                value => TaxonomyAssetApprovalFactory.GroupKey(value.SourceTenantId, value.SourceTermStoreId),
                StringComparer.Ordinal);
            foreach (var pair in reviewedGroupProbes.OrderBy(value => value.Key, StringComparer.Ordinal))
            {
                var reviewed = pair.Value;
                var actionId = TaxonomyAssetApprovalFactory.TermGroupActionId(
                    reviewed.SourceTenantId,
                    reviewed.SourceTermStoreId);
                var decision = approvals[actionId];
                RouteDecision(result, decision);
                if (decision.Decision != TaxonomyAssetApprovalDecision.Approve)
                {
                    continue;
                }
                if (!freshGroupProbes.TryGetValue(pair.Key, out var fresh))
                {
                    AddFailure(result, "TaxonomyTermGroupFreshProbeMissing", actionId, "Fresh inspection omitted the approved TermGroup action.");
                    continue;
                }
                ValidateGroupTransition(decision, reviewed, fresh, result);
            }

            var reviewedSetProbes = reviewedPlan.TermSetProbes.ToDictionary(
                value => TaxonomyAssetApprovalFactory.SetKey(value.SourceTermStoreId, value.SourceTermSetId),
                StringComparer.Ordinal);
            var freshSetProbes = freshInspection.TermSetProbes.ToDictionary(
                value => TaxonomyAssetApprovalFactory.SetKey(value.SourceTermStoreId, value.SourceTermSetId),
                StringComparer.Ordinal);
            foreach (var pair in reviewedSetProbes.OrderBy(value => value.Key, StringComparer.Ordinal))
            {
                var reviewed = pair.Value;
                var actionId = TaxonomyAssetApprovalFactory.TermSetActionId(
                    reviewed.SourceTermStoreId,
                    reviewed.SourceTermSetId);
                var decision = approvals[actionId];
                RouteDecision(result, decision);
                if (decision.Decision != TaxonomyAssetApprovalDecision.Approve)
                {
                    continue;
                }
                if (!freshSetProbes.TryGetValue(pair.Key, out var fresh))
                {
                    AddFailure(result, "TaxonomyTermSetFreshProbeMissing", actionId, "Fresh inspection omitted the approved TermSet action.");
                    continue;
                }
                ValidateSetTransition(decision, reviewed, fresh, result);
            }

            var reviewedTermProbes = reviewedPlan.TermProbes.ToDictionary(
                value => TaxonomyAssetApprovalFactory.TermKey(value.SourceTermStoreId, value.SourceTermSetId, value.SourceTermId),
                StringComparer.Ordinal);
            var freshTermProbes = freshInspection.TermProbes.ToDictionary(
                value => TaxonomyAssetApprovalFactory.TermKey(value.SourceTermStoreId, value.SourceTermSetId, value.SourceTermId),
                StringComparer.Ordinal);
            foreach (var pair in reviewedTermProbes.OrderBy(value => value.Key, StringComparer.Ordinal))
            {
                var reviewed = pair.Value;
                var actionId = TaxonomyAssetApprovalFactory.TermActionId(
                    reviewed.SourceTermStoreId,
                    reviewed.SourceTermSetId,
                    reviewed.SourceTermId);
                var decision = approvals[actionId];
                RouteDecision(result, decision);
                if (decision.Decision != TaxonomyAssetApprovalDecision.Approve)
                {
                    continue;
                }
                if (!freshTermProbes.TryGetValue(pair.Key, out var fresh))
                {
                    AddFailure(result, "TaxonomyTermFreshProbeMissing", actionId, "Fresh inspection omitted the approved Term action.");
                    continue;
                }
                ValidateTermTransition(decision, reviewed, fresh, result);
            }

            result.ApprovedActionIds = result.ApprovedActionIds.Distinct(StringComparer.Ordinal).OrderBy(value => value, StringComparer.Ordinal).ToList();
            result.DeferredActionIds = result.DeferredActionIds.Distinct(StringComparer.Ordinal).OrderBy(value => value, StringComparer.Ordinal).ToList();
            result.RejectedActionIds = result.RejectedActionIds.Distinct(StringComparer.Ordinal).OrderBy(value => value, StringComparer.Ordinal).ToList();
            result.IsAdmitted = result.Failures.Count == 0 && result.ApprovedActionIds.Count > 0;
            return result;
        }

        private static void ValidateGroupTransition(
            TaxonomyAssetActionApproval approval,
            TaxonomyTermGroupTargetProbe reviewed,
            TaxonomyTermGroupTargetProbe fresh,
            TaxonomyAssetExecutionAdmission result)
        {
            var actionId = approval.ActionId;
            var targetId = fresh.ResolvedTargetGroupId ?? approval.TargetTermGroupId;
            if (fresh.SourceTenantId != approval.SourceTenantId
                || fresh.SourceTermStoreId != approval.SourceTermStoreId
                || fresh.TargetTermStoreId != approval.TargetTermStoreId
                || targetId != approval.TargetTermGroupId
                || !IsSafeTransition(reviewed.Disposition, fresh.Disposition))
            {
                AddFailure(result, "TaxonomyTermGroupTargetDrift", actionId,
                    "Fresh TermGroup disposition or target identity differs from the reviewed action: reviewed "
                    + reviewed.Disposition + ", fresh " + fresh.Disposition + ".");
            }
        }

        private static void ValidateSetTransition(
            TaxonomyAssetActionApproval approval,
            TaxonomyTermSetTargetProbe reviewed,
            TaxonomyTermSetTargetProbe fresh,
            TaxonomyAssetExecutionAdmission result)
        {
            var actionId = approval.ActionId;
            var targetId = fresh.ResolvedTargetTermSetId ?? approval.TargetTermSetId;
            if (fresh.TargetTermStoreId != approval.TargetTermStoreId
                || targetId != approval.TargetTermSetId
                || !IsSafeTransition(reviewed.Disposition, fresh.Disposition))
            {
                AddFailure(result, "TaxonomyTermSetTargetDrift", actionId,
                    "Fresh TermSet disposition or target identity differs from the reviewed action: reviewed "
                    + reviewed.Disposition + ", fresh " + fresh.Disposition + ".");
            }
        }

        private static void ValidateTermTransition(
            TaxonomyAssetActionApproval approval,
            TaxonomyTermTargetProbe reviewed,
            TaxonomyTermTargetProbe fresh,
            TaxonomyAssetExecutionAdmission result)
        {
            var actionId = approval.ActionId;
            var targetId = fresh.ResolvedTargetTermId ?? approval.TargetTermId;
            if (fresh.TargetTermStoreId != approval.TargetTermStoreId
                || fresh.TargetTermSetId != approval.TargetTermSetId
                || targetId != approval.TargetTermId
                || !IsSafeTransition(reviewed.Disposition, fresh.Disposition))
            {
                AddFailure(result, "TaxonomyTermTargetDrift", actionId,
                    "Fresh Term disposition or target identity differs from the reviewed action: reviewed "
                    + reviewed.Disposition + ", fresh " + fresh.Disposition + ".");
            }
            if ((reviewed.Disposition == TaxonomyAssetTargetDisposition.ReuseOwned
                    || reviewed.Disposition == TaxonomyAssetTargetDisposition.ReviewExternalReuse)
                && !SameTermRelationshipEvidence(reviewed, fresh))
            {
                AddFailure(result, "TaxonomyTermRelationshipDrift", actionId,
                    "Fresh Term reuse/source-Term or TermSet-membership evidence differs from the reviewed action.");
            }
        }

        private static bool SameTermRelationshipEvidence(
            TaxonomyTermTargetProbe reviewed,
            TaxonomyTermTargetProbe fresh)
        {
            return reviewed.ExistingTermSetId == fresh.ExistingTermSetId
                && reviewed.ExistingIsReused == fresh.ExistingIsReused
                && reviewed.ExistingIsSourceTerm == fresh.ExistingIsSourceTerm
                && reviewed.ExistingReuseSourceTermId == fresh.ExistingReuseSourceTermId
                && reviewed.ExistingPinSourceTermSetId == fresh.ExistingPinSourceTermSetId
                && (reviewed.ExistingTermSetIds ?? new List<Guid>())
                    .Where(value => value != Guid.Empty)
                    .Distinct()
                    .OrderBy(value => value)
                    .SequenceEqual((fresh.ExistingTermSetIds ?? new List<Guid>())
                        .Where(value => value != Guid.Empty)
                        .Distinct()
                        .OrderBy(value => value));
        }

        internal static bool IsSafeTransition(
            TaxonomyAssetTargetDisposition reviewed,
            TaxonomyAssetTargetDisposition fresh)
        {
            if (reviewed == TaxonomyAssetTargetDisposition.ReuseOwned)
            {
                return fresh == TaxonomyAssetTargetDisposition.ReuseOwned;
            }
            if (reviewed == TaxonomyAssetTargetDisposition.CreateMissing)
            {
                return fresh == TaxonomyAssetTargetDisposition.CreateMissing
                    || fresh == TaxonomyAssetTargetDisposition.ReuseOwned;
            }
            if (reviewed == TaxonomyAssetTargetDisposition.ReconcileOwnedPlanDrift)
            {
                return fresh == TaxonomyAssetTargetDisposition.ReconcileOwnedPlanDrift
                    || fresh == TaxonomyAssetTargetDisposition.ReuseOwned;
            }
            if (reviewed == TaxonomyAssetTargetDisposition.ReviewExternalReuse)
            {
                return fresh == TaxonomyAssetTargetDisposition.ReviewExternalReuse;
            }
            if (reviewed == TaxonomyAssetTargetDisposition.CreateMissingAfterExternalApproval)
            {
                return fresh == TaxonomyAssetTargetDisposition.CreateMissingAfterExternalApproval
                    || fresh == TaxonomyAssetTargetDisposition.ReuseOwned;
            }
            return false;
        }

        private static void RouteDecision(
            TaxonomyAssetExecutionAdmission result,
            TaxonomyAssetActionApproval approval)
        {
            if (approval.Decision == TaxonomyAssetApprovalDecision.Approve)
            {
                result.ApprovedActionIds.Add(approval.ActionId);
            }
            else if (approval.Decision == TaxonomyAssetApprovalDecision.Defer)
            {
                result.DeferredActionIds.Add(approval.ActionId);
            }
            else if (approval.Decision == TaxonomyAssetApprovalDecision.Reject)
            {
                result.RejectedActionIds.Add(approval.ActionId);
            }
        }

        private static void AddFailure(
            TaxonomyAssetExecutionAdmission result,
            string code,
            string subject,
            string message)
        {
            result.Failures.Add(new ExecutionAdmissionFailure
            {
                Code = code,
                Subject = subject,
                Message = message
            });
        }
    }
}
