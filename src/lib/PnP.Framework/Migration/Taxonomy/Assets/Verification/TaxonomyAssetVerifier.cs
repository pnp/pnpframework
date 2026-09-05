using PnP.Framework.Migration.Taxonomy.Assets.Execution;
using PnP.Framework.Migration.Taxonomy.Assets.Packaging;
using System;
using System.Collections.Generic;
using System.Linq;

namespace PnP.Framework.Migration.Taxonomy.Assets.Verification
{
    internal static class TaxonomyAssetVerifier
    {
        public static void Verify(
            TaxonomyAssetReviewPlan reviewedPlan,
            TaxonomyAssetApprovalManifest approval,
            TaxonomyAssetReviewPlan freshInspection,
            TaxonomyAssetMaterializationReceipt receipt)
        {
            TaxonomyAssetReviewPlanValidator.Validate(reviewedPlan, true, true);
            TaxonomyAssetApprovalValidator.Validate(reviewedPlan, approval, true, true);
            TaxonomyAssetReviewPlanValidator.Validate(freshInspection, true, true);
            if (receipt == null)
            {
                throw new ArgumentNullException(nameof(receipt));
            }

            var approved = approval.Actions
                .Where(value => value.Decision == TaxonomyAssetApprovalDecision.Approve)
                .ToDictionary(value => value.ActionId, StringComparer.Ordinal);
            var receipts = receipt.Actions.ToDictionary(value => value.ActionId, StringComparer.Ordinal);
            if (approved.Count != receipts.Count
                || approved.Keys.Except(receipts.Keys, StringComparer.Ordinal).Any())
            {
                throw new InvalidOperationException("The taxonomy receipt does not cover every approved action exactly once.");
            }

            var groupProbes = freshInspection.TermGroupProbes.ToDictionary(
                value => TaxonomyAssetApprovalFactory.TermGroupActionId(value.SourceTenantId, value.SourceTermStoreId),
                StringComparer.Ordinal);
            var setProbes = freshInspection.TermSetProbes.ToDictionary(
                value => TaxonomyAssetApprovalFactory.TermSetActionId(value.SourceTermStoreId, value.SourceTermSetId),
                StringComparer.Ordinal);
            var termProbes = freshInspection.TermProbes.ToDictionary(
                value => TaxonomyAssetApprovalFactory.TermActionId(value.SourceTermStoreId, value.SourceTermSetId, value.SourceTermId),
                StringComparer.Ordinal);
            var termPlans = reviewedPlan.Terms.ToDictionary(
                value => TaxonomyAssetApprovalFactory.TermActionId(
                    value.Source.TermStoreId,
                    value.Source.TermSetId,
                    value.Source.TermId),
                StringComparer.Ordinal);
            foreach (var actionReceipt in receipt.Actions)
            {
                var action = approved[actionReceipt.ActionId];
                var expectedFinal = ExpectedFinalDisposition(action.ReviewedDisposition);
                TaxonomyAssetTargetDisposition actualFinal;
                Guid? actualTargetGroup;
                Guid actualTargetSet;
                Guid? actualTargetTerm;
                if (action.Kind == TaxonomyAssetKind.TermGroup)
                {
                    if (!groupProbes.TryGetValue(action.ActionId, out var probe))
                    {
                        throw new InvalidOperationException("Fresh readback omitted approved TermGroup action '" + action.ActionId + "'.");
                    }
                    actualFinal = probe.Disposition;
                    actualTargetGroup = probe.ResolvedTargetGroupId ?? action.TargetTermGroupId;
                    actualTargetSet = Guid.Empty;
                    actualTargetTerm = null;
                }
                else if (action.Kind == TaxonomyAssetKind.TermSet)
                {
                    if (!setProbes.TryGetValue(action.ActionId, out var probe))
                    {
                        throw new InvalidOperationException("Fresh readback omitted approved TermSet action '" + action.ActionId + "'.");
                    }
                    actualFinal = probe.Disposition;
                    actualTargetGroup = action.TargetTermGroupId;
                    actualTargetSet = probe.ResolvedTargetTermSetId ?? action.TargetTermSetId;
                    actualTargetTerm = null;
                }
                else if (action.Kind == TaxonomyAssetKind.Term)
                {
                    if (!termProbes.TryGetValue(action.ActionId, out var probe))
                    {
                        throw new InvalidOperationException("Fresh readback omitted approved Term action '" + action.ActionId + "'.");
                    }
                    actualFinal = probe.Disposition;
                    actualTargetGroup = null;
                    actualTargetSet = probe.TargetTermSetId;
                    actualTargetTerm = probe.ResolvedTargetTermId ?? action.TargetTermId;
                    if (!termPlans.TryGetValue(action.ActionId, out var termPlan))
                    {
                        throw new InvalidOperationException(
                            "Fresh taxonomy Term relationship readback differs for '"
                            + action.ActionId + "': the reviewed Term plan is missing.");
                    }
                    if (!TaxonomyTermRelationshipFidelity.Matches(
                            termPlan,
                            probe,
                            action.TargetTermSetId,
                            out var relationshipDiagnostic))
                    {
                        throw new InvalidOperationException(
                            "Fresh taxonomy Term relationship readback differs for '"
                            + action.ActionId + "': " + relationshipDiagnostic);
                    }
                }
                else
                {
                    throw new InvalidOperationException("Unsupported taxonomy action kind in receipt: " + action.Kind + ".");
                }

                if (actualFinal != expectedFinal
                    || actualTargetGroup != action.TargetTermGroupId
                    || actualTargetSet != action.TargetTermSetId
                    || actualTargetTerm != action.TargetTermId)
                {
                    throw new InvalidOperationException(
                        "Fresh taxonomy readback differs for '" + action.ActionId + "': expected "
                        + expectedFinal + ", observed " + actualFinal + ".");
                }
                actionReceipt.FinalDisposition = actualFinal;
                actionReceipt.FreshReadbackPassed = true;
                actionReceipt.Diagnostic = "Fresh target inspection matched the approved taxonomy identity, shape, and captured relationship evidence.";
            }

            receipt.ChangedTarget = receipt.Actions.Any(value => value.ChangedTarget);
            receipt.FreshReadbackPassed = true;
            receipt.Diagnostics.Add(
                "Fresh target inspection verified " + receipt.Actions.Count
                + " approved taxonomy asset action(s); deferred and rejected actions made no target claim.");
        }

        internal static TaxonomyAssetTargetDisposition ExpectedFinalDisposition(
            TaxonomyAssetTargetDisposition reviewedDisposition)
        {
            if (reviewedDisposition == TaxonomyAssetTargetDisposition.ReviewExternalReuse)
            {
                return TaxonomyAssetTargetDisposition.ReviewExternalReuse;
            }
            if (reviewedDisposition == TaxonomyAssetTargetDisposition.CreateMissing
                || reviewedDisposition == TaxonomyAssetTargetDisposition.ReuseOwned
                || reviewedDisposition == TaxonomyAssetTargetDisposition.ReconcileOwnedPlanDrift
                || reviewedDisposition == TaxonomyAssetTargetDisposition.CreateMissingAfterExternalApproval)
            {
                return TaxonomyAssetTargetDisposition.ReuseOwned;
            }
            throw new InvalidOperationException("Reviewed taxonomy disposition is not executable: " + reviewedDisposition + ".");
        }
    }
}
