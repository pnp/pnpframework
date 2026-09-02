using Microsoft.SharePoint.Client;
using PnP.Framework.Migration.Execution;
using PnP.Framework.Migration.Pages;
using PnP.Framework.Migration.Pages.Publishing.Execution;
using PnP.Framework.Migration.Pages.Publishing.Packaging;
using PnP.Framework.Migration.Pages.Publishing.Layouts;
using PnP.Framework.Migration.Pages.Publishing.Verification;
using System;
using System.Collections.Generic;
using System.Linq;
using PnP.Framework.Migration.Topology;
using PnP.Framework.Migration.Lists.Planning;

namespace PnP.Framework.Migration.Pages.Publishing.Execution
{
    internal static class PublishingPageImportAdmission
    {
        public static PublishingPageImportReceipt TryAdmit(
            ClientContext targetContext,
            PublishingPageMigrationPackage package,
            string approvedPlanDigest,
            Guid operationId,
            DateTimeOffset startedAt,
            MigrationExecutionRecorder recorder)
        {
            if (package.State != PublishingPagePackageState.ApprovalReady || !package.Plan.IsExecutable)
            {
                return Failure(package, operationId, startedAt, "PlanNotExecutable", package.Plan.TargetPageServerRelativeUrl,
                    "The publishing-page package is not approval-ready.", recorder);
            }

            if (string.IsNullOrWhiteSpace(approvedPlanDigest)
                || !string.Equals(approvedPlanDigest, package.PlanDigest, StringComparison.OrdinalIgnoreCase))
            {
                return Failure(package, operationId, startedAt, "PlanDigestNotApproved", package.Plan.TargetPageServerRelativeUrl,
                    "The approved plan digest does not match the sealed publishing-page package.", recorder);
            }

            var targetWeb = targetContext.Web;
            targetContext.Load(targetWeb, web => web.Url, web => web.ServerRelativeUrl);
            targetContext.ExecuteQueryRetry();
            if (!PagePath.UriEquals(targetWeb.Url, package.Plan.TargetWebUrl))
            {
                return Failure(package, operationId, startedAt, "TargetIdentityMismatch", targetWeb.Url,
                    $"The target connection points to '{targetWeb.Url}', but the approved plan targets '{package.Plan.TargetWebUrl}'.", recorder);
            }

            var blockers = new List<string>();
            TopologyTargetAnalysis freshTopology = null;
            if (package.Plan.Topology != null)
            {
                freshTopology = TopologyTargetInspector.Inspect(targetContext, package.Plan.Topology, targetWeb.Url);
                blockers.AddRange(freshTopology.Issues.Select(value => value.Code + ": " + value.Message));
            }
            if (package.Plan.ListMigration != null)
            {
                var freshLists = ListMigrationTargetAnalyzer.InspectFresh(
                    targetContext,
                    package.Snapshot.ListDependencies,
                    package.Plan.ListMigration,
                    freshTopology);
                blockers.AddRange(freshLists.Issues.Select(value => value.Code + ": " + value.Message));
            }
            var freshLayoutProbe = PublishingPageLayoutTargetInspector.Inspect(
                targetContext,
                package.Plan.LayoutMaterialization);
            var freshLayoutAdmission = PublishingPageLayoutTargetAdmissionEvaluator.Evaluate(
                package.Plan.LayoutMaterialization,
                freshLayoutProbe);
            foreach (var issue in freshLayoutAdmission.Issues)
            {
                blockers.Add($"{issue.Code}: {issue.Message}");
            }

            var freshProbe = PublishingPageTargetInspector.Inspect(
                targetContext,
                package.Plan.TargetPageServerRelativeUrl,
                package.Plan.DependencyActions,
                package.Plan.TargetLifecycle,
                package.Plan.LayoutMaterialization,
                freshLayoutProbe,
                blockers);
            if (blockers.Count > 0)
            {
                return Failure(package, operationId, startedAt, DetermineFailureCode(freshProbe), package.Plan.TargetPageServerRelativeUrl,
                    "Fresh target preflight failed: " + string.Join(" ", blockers), recorder);
            }

            if (!string.Equals(freshProbe.PageContentTypeId, package.Plan.TargetProbe.PageContentTypeId, StringComparison.OrdinalIgnoreCase)
                || !string.Equals(freshProbe.PageLayoutUrl, package.Plan.TargetProbe.PageLayoutUrl, StringComparison.OrdinalIgnoreCase))
            {
                return Failure(package, operationId, startedAt, "TargetPreconditionChanged", package.Plan.TargetPageServerRelativeUrl,
                    "The target Publishing Page content type or layout changed after approval.", recorder);
            }

            return null;
        }

        private static PublishingPageImportReceipt Failure(
            PublishingPageMigrationPackage package,
            Guid operationId,
            DateTimeOffset startedAt,
            string code,
            string subject,
            string message,
            MigrationExecutionRecorder recorder)
        {
            return PublishingPageImportReceiptFactory.AdmissionFailure(
                package,
                operationId,
                startedAt,
                code,
                subject,
                message,
                recorder);
        }

        private static string DetermineFailureCode(PublishingPageTargetSnapshot probe)
        {
            if (probe.TargetPageExists)
            {
                return "CreateOnlyTargetExists";
            }

            if (probe.ExistingDependencyPaths.Count > 0)
            {
                return "DependencyTargetCollision";
            }

            return "TargetPreflightFailed";
        }
    }
}
