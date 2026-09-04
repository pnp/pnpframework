using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.Taxonomy;
using PnP.Framework.Migration.Execution;
using PnP.Framework.Migration.Taxonomy.Assets.Packaging;
using PnP.Framework.Migration.Taxonomy.Assets.Verification;
using System;
using System.Collections.Generic;
using System.Linq;

namespace PnP.Framework.Migration.Taxonomy.Assets.Execution
{
    internal static class TaxonomyAssetMaterializationCoordinator
    {
        public static TaxonomyAssetMigrationExecutionResult Ensure(
            ClientContext context,
            TaxonomyAssetReviewPlan reviewedPlan,
            TaxonomyAssetApprovalManifest approval,
            MigrationExecutionRecorder recorder)
        {
            var freshPreflight = InspectClone(context, reviewedPlan);
            var admission = TaxonomyAssetExecutionAdmissionEvaluator.Evaluate(
                reviewedPlan,
                freshPreflight,
                approval);
            if (!admission.IsAdmitted)
            {
                throw new InvalidOperationException(
                    "Fresh taxonomy asset admission failed: "
                    + string.Join("; ", admission.Failures.Select(value => value.Message)));
            }

            var receipt = new TaxonomyAssetMaterializationReceipt
            {
                OperationId = recorder.OperationId,
                ReviewPlanDigest = reviewedPlan.PlanDigest,
                ApprovalDigest = approval.ApprovalDigest,
                TargetTermStoreId = reviewedPlan.TargetTermStoreId,
                StartedAtUtc = DateTimeOffset.UtcNow,
                DeferredActionIds = admission.DeferredActionIds.ToList(),
                RejectedActionIds = admission.RejectedActionIds.ToList()
            };
            var approved = approval.Actions
                .Where(value => value.Decision == TaxonomyAssetApprovalDecision.Approve)
                .ToDictionary(value => value.ActionId, StringComparer.Ordinal);
            var store = TaxonomyAssetCsomMaterializer.GetStore(context, reviewedPlan.TargetTermStoreId);
            EnsureTermGroups(context, store, reviewedPlan, freshPreflight, approved, recorder, receipt);
            EnsureTermSets(context, store, reviewedPlan, freshPreflight, approved, recorder, receipt);
            EnsureTerms(context, store, reviewedPlan, freshPreflight, approved, recorder, receipt);

            var finalInspection = InspectClone(context, reviewedPlan);
            TaxonomyAssetVerifier.Verify(reviewedPlan, approval, finalInspection, receipt);
            receipt.CompletedAtUtc = DateTimeOffset.UtcNow;
            TaxonomyAssetMaterializationReceiptValidator.Seal(reviewedPlan, approval, receipt);
            return new TaxonomyAssetMigrationExecutionResult
            {
                OperationId = recorder.OperationId,
                Admission = admission,
                Receipt = receipt,
                Steps = recorder.Steps.ToList()
            };
        }

        private static void EnsureTermGroups(
            ClientContext context,
            TermStore store,
            TaxonomyAssetReviewPlan reviewedPlan,
            TaxonomyAssetReviewPlan freshPreflight,
            IDictionary<string, TaxonomyAssetActionApproval> approved,
            MigrationExecutionRecorder recorder,
            TaxonomyAssetMaterializationReceipt receipt)
        {
            var fresh = freshPreflight.TermGroupProbes.ToDictionary(
                value => TaxonomyAssetApprovalFactory.GroupKey(value.SourceTenantId, value.SourceTermStoreId),
                StringComparer.Ordinal);
            foreach (var plan in reviewedPlan.TermGroups
                         .OrderBy(value => value.Source.TenantId)
                         .ThenBy(value => value.Source.TermStoreId))
            {
                var actionId = TaxonomyAssetApprovalFactory.TermGroupActionId(
                    plan.Source.TenantId,
                    plan.Source.TermStoreId);
                if (!approved.TryGetValue(actionId, out var action))
                {
                    continue;
                }
                var probe = fresh[TaxonomyAssetApprovalFactory.GroupKey(
                    plan.Source.TenantId,
                    plan.Source.TermStoreId)];
                var changed = false;
                if (probe.Disposition == TaxonomyAssetTargetDisposition.ReuseOwned)
                {
                    recorder.RecordAlreadySatisfied(actionId, "Fresh preflight approved reusable taxonomy TermGroup '" + plan.TargetGroupName + "'.");
                }
                else
                {
                    changed = recorder.Execute(
                        actionId,
                        "Ensure taxonomy TermGroup '" + plan.TargetGroupName + "'.",
                        () => TaxonomyAssetCsomMaterializer.EnsureOwnedGroup(
                            context,
                            store,
                            plan.PreferredTargetGroupId,
                            plan.TargetGroupName),
                        value => value ? MutationOutcome.Applied : MutationOutcome.AlreadySatisfied,
                        value => value ? "Created and freshly verified taxonomy TermGroup." : "Reused exact taxonomy TermGroup.");
                }
                receipt.Actions.Add(GroupReceipt(action, probe, changed));
            }
        }

        private static void EnsureTermSets(
            ClientContext context,
            TermStore store,
            TaxonomyAssetReviewPlan reviewedPlan,
            TaxonomyAssetReviewPlan freshPreflight,
            IDictionary<string, TaxonomyAssetActionApproval> approved,
            MigrationExecutionRecorder recorder,
            TaxonomyAssetMaterializationReceipt receipt)
        {
            var fresh = freshPreflight.TermSetProbes.ToDictionary(
                value => TaxonomyAssetApprovalFactory.SetKey(value.SourceTermStoreId, value.SourceTermSetId),
                StringComparer.Ordinal);
            foreach (var plan in reviewedPlan.TermSets
                         .OrderBy(value => value.Source.TermStoreId)
                         .ThenBy(value => value.Source.TermSetId))
            {
                var actionId = TaxonomyAssetApprovalFactory.TermSetActionId(
                    plan.Source.TermStoreId,
                    plan.Source.TermSetId);
                if (!approved.TryGetValue(actionId, out var action))
                {
                    continue;
                }
                var probe = fresh[TaxonomyAssetApprovalFactory.SetKey(
                    plan.Source.TermStoreId,
                    plan.Source.TermSetId)];
                var changed = false;
                if (probe.Disposition == TaxonomyAssetTargetDisposition.ReuseOwned
                    || probe.Disposition == TaxonomyAssetTargetDisposition.ReviewExternalReuse)
                {
                    recorder.RecordAlreadySatisfied(actionId, "Fresh preflight approved reusable TermSet '" + plan.SourceTermSetName + "'.");
                }
                else
                {
                    changed = recorder.Execute(
                        actionId,
                        "Ensure TermSet '" + plan.SourceTermSetName + "'.",
                        () => TaxonomyAssetCsomMaterializer.EnsureTermSet(context, store, plan, probe),
                        value => value ? MutationOutcome.Applied : MutationOutcome.AlreadySatisfied,
                        value => value ? "TermSet mutation committed; aggregate fresh verification is pending." : "TermSet already matched.");
                }
                receipt.Actions.Add(SetReceipt(action, probe, changed));
            }
        }

        private static void EnsureTerms(
            ClientContext context,
            TermStore store,
            TaxonomyAssetReviewPlan reviewedPlan,
            TaxonomyAssetReviewPlan freshPreflight,
            IDictionary<string, TaxonomyAssetActionApproval> approved,
            MigrationExecutionRecorder recorder,
            TaxonomyAssetMaterializationReceipt receipt)
        {
            var fresh = freshPreflight.TermProbes.ToDictionary(
                value => TaxonomyAssetApprovalFactory.TermKey(
                    value.SourceTermStoreId,
                    value.SourceTermSetId,
                    value.SourceTermId),
                StringComparer.Ordinal);
            foreach (var plan in OrderTerms(reviewedPlan.Terms))
            {
                var actionId = TaxonomyAssetApprovalFactory.TermActionId(
                    plan.Source.TermStoreId,
                    plan.Source.TermSetId,
                    plan.Source.TermId);
                if (!approved.TryGetValue(actionId, out var action))
                {
                    continue;
                }
                var probe = fresh[TaxonomyAssetApprovalFactory.TermKey(
                    plan.Source.TermStoreId,
                    plan.Source.TermSetId,
                    plan.Source.TermId)];
                var changed = false;
                if (probe.Disposition == TaxonomyAssetTargetDisposition.ReuseOwned
                    || probe.Disposition == TaxonomyAssetTargetDisposition.ReviewExternalReuse)
                {
                    recorder.RecordAlreadySatisfied(actionId, "Fresh preflight approved reusable Term '" + plan.Name + "'.");
                }
                else
                {
                    changed = recorder.Execute(
                        actionId,
                        "Ensure Term '" + plan.Name + "'.",
                        () => TaxonomyAssetCsomMaterializer.EnsureTerm(context, store, plan, probe),
                        value => value ? MutationOutcome.Applied : MutationOutcome.AlreadySatisfied,
                        value => value ? "Term mutation committed; aggregate fresh verification is pending." : "Term already matched.");
                }
                receipt.Actions.Add(TermReceipt(action, probe, changed));
            }
        }

        private static TaxonomyAssetReviewPlan InspectClone(
            ClientContext context,
            TaxonomyAssetReviewPlan reviewedPlan)
        {
            var clone = TaxonomyAssetContractCloner.Clone(reviewedPlan);
            return TaxonomyAssetTargetInspector.Inspect(context, clone);
        }

        internal static IEnumerable<TaxonomyTermMaterializationPlan> OrderTerms(
            IEnumerable<TaxonomyTermMaterializationPlan> values)
        {
            var remaining = (values ?? Enumerable.Empty<TaxonomyTermMaterializationPlan>())
                .Where(value => value != null && value.Source != null)
                .ToList();
            var emitted = new HashSet<string>(StringComparer.Ordinal);
            while (remaining.Count > 0)
            {
                var ready = remaining
                    .Where(value => !value.TargetParentTermId.HasValue
                        || emitted.Contains(TaxonomyAssetApprovalFactory.TermKey(
                            value.Source.TermStoreId,
                            value.Source.TermSetId,
                            value.TargetParentTermId.Value)))
                    .OrderBy(value => value.Source.TermStoreId)
                    .ThenBy(value => value.Source.TermSetId)
                    .ThenBy(value => value.Source.TermId)
                    .ToArray();
                if (ready.Length == 0)
                {
                    throw new InvalidOperationException("The approved taxonomy Term closure has a cyclic or missing parent dependency.");
                }
                foreach (var item in ready)
                {
                    remaining.Remove(item);
                    emitted.Add(TaxonomyAssetApprovalFactory.TermKey(
                        item.Source.TermStoreId,
                        item.Source.TermSetId,
                        item.Source.TermId));
                    yield return item;
                }
            }
        }

        private static TaxonomyAssetActionReceipt GroupReceipt(
            TaxonomyAssetActionApproval action,
            TaxonomyTermGroupTargetProbe probe,
            bool changed)
        {
            return new TaxonomyAssetActionReceipt
            {
                ActionId = action.ActionId,
                Kind = TaxonomyAssetKind.TermGroup,
                SourceTenantId = action.SourceTenantId,
                SourceTermStoreId = action.SourceTermStoreId,
                TargetTermStoreId = action.TargetTermStoreId,
                TargetTermGroupId = action.TargetTermGroupId,
                ReviewedDisposition = action.ReviewedDisposition,
                PreflightDisposition = probe.Disposition,
                ChangedTarget = changed
            };
        }

        private static TaxonomyAssetActionReceipt SetReceipt(
            TaxonomyAssetActionApproval action,
            TaxonomyTermSetTargetProbe probe,
            bool changed)
        {
            return new TaxonomyAssetActionReceipt
            {
                ActionId = action.ActionId,
                Kind = TaxonomyAssetKind.TermSet,
                SourceTenantId = action.SourceTenantId,
                SourceTermStoreId = action.SourceTermStoreId,
                SourceTermSetId = action.SourceTermSetId,
                TargetTermStoreId = action.TargetTermStoreId,
                TargetTermGroupId = action.TargetTermGroupId,
                TargetTermSetId = action.TargetTermSetId,
                ReviewedDisposition = action.ReviewedDisposition,
                PreflightDisposition = probe.Disposition,
                ChangedTarget = changed
            };
        }

        private static TaxonomyAssetActionReceipt TermReceipt(
            TaxonomyAssetActionApproval action,
            TaxonomyTermTargetProbe probe,
            bool changed)
        {
            return new TaxonomyAssetActionReceipt
            {
                ActionId = action.ActionId,
                Kind = TaxonomyAssetKind.Term,
                SourceTenantId = action.SourceTenantId,
                SourceTermStoreId = action.SourceTermStoreId,
                SourceTermSetId = action.SourceTermSetId,
                SourceTermId = action.SourceTermId,
                TargetTermStoreId = action.TargetTermStoreId,
                TargetTermSetId = action.TargetTermSetId,
                TargetTermId = action.TargetTermId,
                ReviewedDisposition = action.ReviewedDisposition,
                PreflightDisposition = probe.Disposition,
                ChangedTarget = changed
            };
        }
    }
}
