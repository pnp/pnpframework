using PnP.Framework.Migration.Packaging;
using PnP.Framework.Migration.Taxonomy.Assets.Execution;
using PnP.Framework.Migration.Taxonomy.Assets.Verification;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace PnP.Framework.Migration.Taxonomy.Assets.Packaging
{
    public static class TaxonomyAssetMaterializationReceiptValidator
    {
        public static void Seal(
            TaxonomyAssetReviewPlan reviewPlan,
            TaxonomyAssetApprovalManifest approval,
            TaxonomyAssetMaterializationReceipt receipt)
        {
            if (receipt == null)
            {
                throw new ArgumentNullException(nameof(receipt));
            }
            receipt.ReceiptDigest = null;
            Validate(reviewPlan, approval, receipt, false);
            receipt.ReceiptDigest = ComputeDigest(receipt);
            Validate(reviewPlan, approval, receipt, true);
        }

        public static void Validate(
            TaxonomyAssetReviewPlan reviewPlan,
            TaxonomyAssetApprovalManifest approval,
            TaxonomyAssetMaterializationReceipt receipt,
            bool requireDigest = true)
        {
            TaxonomyAssetReviewPlanValidator.Validate(reviewPlan, true, true);
            TaxonomyAssetApprovalValidator.Validate(reviewPlan, approval, true, true);
            if (receipt == null)
            {
                throw new ArgumentNullException(nameof(receipt));
            }

            var errors = new List<string>();
            if (!string.Equals(receipt.SchemaVersion, "pnp-taxonomy-asset-materialization-receipt/v1", StringComparison.Ordinal))
            {
                errors.Add("Unsupported taxonomy asset materialization-receipt schema.");
            }
            if (receipt.OperationId == Guid.Empty)
            {
                errors.Add("The taxonomy materialization operation identity is missing.");
            }
            if (!string.Equals(receipt.ReviewPlanDigest, reviewPlan.PlanDigest, StringComparison.OrdinalIgnoreCase)
                || !string.Equals(receipt.ApprovalDigest, approval.ApprovalDigest, StringComparison.OrdinalIgnoreCase)
                || receipt.TargetTermStoreId != reviewPlan.TargetTermStoreId)
            {
                errors.Add("The taxonomy receipt does not bind the reviewed plan, approval, and target TermStore boundary.");
            }
            if (receipt.StartedAtUtc == default(DateTimeOffset)
                || receipt.CompletedAtUtc == default(DateTimeOffset)
                || receipt.CompletedAtUtc < receipt.StartedAtUtc)
            {
                errors.Add("The taxonomy receipt has invalid start/completion timestamps.");
            }
            if (!receipt.FreshReadbackPassed)
            {
                errors.Add("The taxonomy receipt does not prove aggregate fresh readback.");
            }

            var expectedApproved = approval.Actions
                .Where(value => value.Decision == TaxonomyAssetApprovalDecision.Approve)
                .ToDictionary(value => value.ActionId, StringComparer.Ordinal);
            var actual = new Dictionary<string, TaxonomyAssetActionReceipt>(StringComparer.Ordinal);
            foreach (var action in receipt.Actions ?? new List<TaxonomyAssetActionReceipt>())
            {
                if (action == null || string.IsNullOrWhiteSpace(action.ActionId) || actual.ContainsKey(action.ActionId))
                {
                    errors.Add("The taxonomy receipt contains a null, blank, or duplicate action.");
                    continue;
                }
                actual[action.ActionId] = action;
                if (!expectedApproved.TryGetValue(action.ActionId, out var approved))
                {
                    errors.Add("The taxonomy receipt contains an action that was not approved: '" + action.ActionId + "'.");
                    continue;
                }
                ValidateAction(approved, action, errors);
            }
            foreach (var missing in expectedApproved.Keys.Except(actual.Keys, StringComparer.Ordinal))
            {
                errors.Add("The taxonomy receipt omits approved action '" + missing + "'.");
            }

            var expectedDeferred = approval.Actions
                .Where(value => value.Decision == TaxonomyAssetApprovalDecision.Defer)
                .Select(value => value.ActionId);
            var expectedRejected = approval.Actions
                .Where(value => value.Decision == TaxonomyAssetApprovalDecision.Reject)
                .Select(value => value.ActionId);
            if (!SetEquals(receipt.DeferredActionIds, expectedDeferred)
                || !SetEquals(receipt.RejectedActionIds, expectedRejected))
            {
                errors.Add("The taxonomy receipt deferred/rejected action sets differ from the sealed approval.");
            }
            if (receipt.ChangedTarget != actual.Values.Any(value => value.ChangedTarget))
            {
                errors.Add("The taxonomy receipt aggregate ChangedTarget value differs from its action receipts.");
            }
            if (requireDigest && (!IsSha256(receipt.ReceiptDigest)
                || !string.Equals(receipt.ReceiptDigest, ComputeDigest(receipt), StringComparison.OrdinalIgnoreCase)))
            {
                errors.Add("The taxonomy materialization receipt digest is absent or invalid.");
            }
            if (errors.Count > 0)
            {
                throw new InvalidDataException("Invalid taxonomy asset materialization receipt: " + string.Join(" ", errors));
            }
        }

        public static string ComputeDigest(TaxonomyAssetMaterializationReceipt receipt)
        {
            if (receipt == null)
            {
                throw new ArgumentNullException(nameof(receipt));
            }

            return MigrationDigest.ComputeSha256(
                MigrationContractSerializer.SerializeCanonicalWithNullRootProperty(
                    receipt,
                    nameof(TaxonomyAssetMaterializationReceipt.ReceiptDigest)));
        }

        private static void ValidateAction(
            TaxonomyAssetActionApproval approved,
            TaxonomyAssetActionReceipt actual,
            ICollection<string> errors)
        {
            if (actual.Kind != approved.Kind
                || actual.SourceTenantId != approved.SourceTenantId
                || actual.SourceTermStoreId != approved.SourceTermStoreId
                || actual.SourceTermSetId != approved.SourceTermSetId
                || actual.SourceTermId != approved.SourceTermId
                || actual.TargetTermStoreId != approved.TargetTermStoreId
                || actual.TargetTermGroupId != approved.TargetTermGroupId
                || actual.TargetTermSetId != approved.TargetTermSetId
                || actual.TargetTermId != approved.TargetTermId
                || actual.ReviewedDisposition != approved.ReviewedDisposition)
            {
                errors.Add("Taxonomy receipt action '" + actual.ActionId + "' differs from its approved identity or disposition.");
            }
            if (!TaxonomyAssetExecutionAdmissionEvaluator.IsSafeTransition(
                    approved.ReviewedDisposition,
                    actual.PreflightDisposition)
                || actual.FinalDisposition != TaxonomyAssetVerifier.ExpectedFinalDisposition(approved.ReviewedDisposition)
                || !actual.FreshReadbackPassed)
            {
                errors.Add("Taxonomy receipt action '" + actual.ActionId + "' lacks a safe preflight transition or exact fresh readback result.");
            }
        }

        private static bool SetEquals(IEnumerable<string> left, IEnumerable<string> right)
        {
            return new HashSet<string>(left ?? Enumerable.Empty<string>(), StringComparer.Ordinal)
                .SetEquals(right ?? Enumerable.Empty<string>());
        }

        private static bool IsSha256(string value)
        {
            return !string.IsNullOrWhiteSpace(value)
                && value.Length == 64
                && value.All(character => character >= '0' && character <= '9'
                    || character >= 'a' && character <= 'f'
                    || character >= 'A' && character <= 'F');
        }
    }
}
