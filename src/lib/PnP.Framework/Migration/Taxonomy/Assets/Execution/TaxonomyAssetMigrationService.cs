using Microsoft.SharePoint.Client;
using PnP.Framework.Migration.Execution;
using PnP.Framework.Migration.Taxonomy.Assets.Packaging;
using System;
using System.Linq;

namespace PnP.Framework.Migration.Taxonomy.Assets.Execution
{
    /// <summary>
    /// Public application boundary for inspecting and materializing an explicitly
    /// approved taxonomy asset closure. Page relationship replay remains a separate
    /// consumer and cannot implicitly create or repair taxonomy assets.
    /// </summary>
    public sealed class TaxonomyAssetMigrationService
    {
        public TaxonomyAssetReviewPlan Inspect(
            ClientContext targetContext,
            TaxonomyAssetReviewPlan plan)
        {
            if (targetContext == null)
            {
                throw new ArgumentNullException(nameof(targetContext));
            }
            TaxonomyAssetReviewPlanValidator.Validate(plan, true, false);
            var clone = TaxonomyAssetContractCloner.Clone(plan);
            return TaxonomyAssetTargetInspector.Inspect(targetContext, clone);
        }

        public TaxonomyAssetMigrationExecutionResult Ensure(
            ClientContext targetContext,
            TaxonomyAssetReviewPlan reviewedPlan,
            TaxonomyAssetApprovalManifest approval,
            IMigrationExecutionJournal journal = null)
        {
            if (targetContext == null)
            {
                throw new ArgumentNullException(nameof(targetContext));
            }
            TaxonomyAssetReviewPlanValidator.Validate(reviewedPlan, true, true);
            TaxonomyAssetApprovalValidator.Validate(reviewedPlan, approval, true, true);

            var operationId = Guid.NewGuid();
            var recorder = new MigrationExecutionRecorder(operationId, reviewedPlan.PlanDigest, journal);
            recorder.RecordState(MigrationExecutionStatus.Running, "Taxonomy asset materialization started.");
            try
            {
                var result = TaxonomyAssetMaterializationCoordinator.Ensure(
                    targetContext,
                    reviewedPlan,
                    approval,
                    recorder);
                result.Steps = recorder.Steps.ToList();
                recorder.RecordState(MigrationExecutionStatus.Succeeded, "Taxonomy asset materialization and fresh readback completed.");
                return result;
            }
            catch
            {
                recorder.RecordState(MigrationExecutionStatus.FailedUnexpectedly, "Taxonomy asset materialization failed.");
                throw;
            }
        }
    }
}
