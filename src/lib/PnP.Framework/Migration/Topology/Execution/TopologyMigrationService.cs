using Microsoft.SharePoint.Client;
using PnP.Framework.Migration.Execution;
using System;
using System.Collections.Generic;

namespace PnP.Framework.Migration.Topology.Execution
{
    public sealed class TopologyMigrationExecutionResult
    {
        public Guid OperationId { get; set; }

        public TopologyMaterializationReceipt Receipt { get; set; }

        public IList<MigrationMutationReceipt> Steps { get; set; } = new List<MigrationMutationReceipt>();
    }

    /// <summary>
    /// Public application boundary for inspecting and materializing an approved Site/Web topology plan.
    /// Site Collection creation remains a tenant-host responsibility; this service creates missing child Webs.
    /// </summary>
    public sealed class TopologyMigrationService
    {
        public TopologyTargetAnalysis Inspect(
            ClientContext targetContext,
            TopologyPlan plan,
            string approvedHostWebUrl = null)
        {
            if (targetContext == null)
            {
                throw new ArgumentNullException(nameof(targetContext));
            }
            if (plan == null)
            {
                throw new ArgumentNullException(nameof(plan));
            }
            TopologyPlanValidator.Validate(plan);
            return TopologyTargetInspector.Inspect(targetContext, plan, approvedHostWebUrl ?? targetContext.Url);
        }

        /// <summary>
        /// Performs the mutable planning phase that resolves proven foreign child-Web
        /// collisions with a deterministic suffix at that Web node and reseals the
        /// topology digest. Approved plans must use <see cref="Inspect"/> instead.
        /// </summary>
        public TopologyTargetAnalysis ResolvePlanningCollisions(
            ClientContext targetContext,
            TopologyPlan plan,
            string approvedHostWebUrl = null)
        {
            if (targetContext == null)
            {
                throw new ArgumentNullException(nameof(targetContext));
            }
            if (plan == null)
            {
                throw new ArgumentNullException(nameof(plan));
            }
            TopologyPlanValidator.Validate(plan);
            return TopologyTargetInspector.InspectForPlanning(
                targetContext,
                plan,
                approvedHostWebUrl ?? targetContext.Url);
        }

        /// <summary>
        /// Moves one Site Collection mapping to a reviewed collision-safe URL while
        /// preserving every Web-relative path below that Site Collection. This is a
        /// planning-only operation and reseals the aggregate topology digest.
        /// </summary>
        public void RetargetSiteCollision(
            TopologyPlan plan,
            Guid sourceSiteId,
            string targetSiteCollectionUrl,
            string reason)
        {
            if (plan == null)
            {
                throw new ArgumentNullException(nameof(plan));
            }
            var sitePlan = System.Linq.Enumerable.Single(
                plan.SiteCollections,
                value => value.SourceSiteId == sourceSiteId);
            TopologyPlanRetargeter.RetargetSiteCollection(sitePlan, targetSiteCollectionUrl, reason);
            plan.PlanDigest = TopologyPlanner.ComputeDigest(plan);
        }

        public TopologyMigrationExecutionResult Ensure(
            ClientContext targetContext,
            TopologyPlan plan,
            IMigrationExecutionJournal journal = null)
        {
            if (targetContext == null)
            {
                throw new ArgumentNullException(nameof(targetContext));
            }
            if (plan == null)
            {
                throw new ArgumentNullException(nameof(plan));
            }
            TopologyPlanValidator.Validate(plan);

            var operationId = Guid.NewGuid();
            var recorder = new MigrationExecutionRecorder(operationId, plan.PlanDigest, journal);
            recorder.RecordState(MigrationExecutionStatus.Running, "Topology materialization started.");
            try
            {
                var receipt = TopologyMaterializationCoordinator.Ensure(targetContext, plan, recorder);
                recorder.RecordState(MigrationExecutionStatus.Succeeded, "Topology materialization and fresh readback completed.");
                return new TopologyMigrationExecutionResult
                {
                    OperationId = operationId,
                    Receipt = receipt,
                    Steps = new List<MigrationMutationReceipt>(recorder.Steps)
                };
            }
            catch
            {
                recorder.RecordState(MigrationExecutionStatus.FailedUnexpectedly, "Topology materialization failed.");
                throw;
            }
        }
    }
}
