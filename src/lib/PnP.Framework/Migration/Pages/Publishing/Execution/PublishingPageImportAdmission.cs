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
using PnP.Framework.Migration.Pages.Fields.Taxonomy;

namespace PnP.Framework.Migration.Pages.Publishing.Execution
{
    internal static class PublishingPageImportAdmission
    {
        public static PublishingPageImportReceipt TryAdmit(
            ClientContext targetContext,
            PublishingPageMigrationPackage package,
            PublishingPageExecutionScope executionScope,
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
            if (executionScope.PageArtifact && !PagePath.UriEquals(targetWeb.Url, package.Plan.TargetWebUrl))
            {
                return Failure(package, operationId, startedAt, "TargetIdentityMismatch", targetWeb.Url,
                    $"The target connection points to '{targetWeb.Url}', but the approved plan targets '{package.Plan.TargetWebUrl}'.", recorder);
            }
            if (!SameAuthority(targetWeb.Url, package.Plan.TargetWebUrl))
            {
                return Failure(package, operationId, startedAt, "TargetTenantMismatch", targetWeb.Url,
                    $"The target connection authority '{new Uri(targetWeb.Url).Authority}' differs from the approved target authority '{new Uri(package.Plan.TargetWebUrl).Authority}'.", recorder);
            }

            var blockers = new List<string>();
            var taxonomyActions = executionScope.TaxonomyActions(package);
            if (taxonomyActions.Count > 0)
            {
                blockers.AddRange(PageTaxonomyRelationshipPlanner.ValidateFresh(
                    targetContext,
                    package.Snapshot.Fields,
                    taxonomyActions,
                    package.Plan.PlanningPolicy));
            }
            TopologyTargetAnalysis freshTopology = null;
            if (executionScope.TopologyPlan != null)
            {
                freshTopology = TopologyTargetInspector.Inspect(targetContext, executionScope.TopologyPlan, targetWeb.Url);
                blockers.AddRange(freshTopology.Issues.Select(value => value.Code + ": " + value.Message));
            }
            if (executionScope.ListScope?.HasWork == true)
            {
                var sourceById = package.Snapshot.ListDependencies.ToDictionary(value => value.SourceListId);
                var listSelections = executionScope.ListScope.Lists
                    .Where(value => value.HasListScopedWork)
                    .ToArray();
                var projectedSources = listSelections
                    .Select(value => executionScope.ListScope.ProjectSource(sourceById[value.SourceListId]))
                    .ToArray();
                if (projectedSources.Length > 0)
                {
                    var selectedListIds = new HashSet<Guid>(
                        listSelections.Select(value => value.SourceListId));
                    var projectedPlan = executionScope.ListScope.ProjectPlanSet(package.Plan.ListMigration);
                    projectedPlan.OrderedSourceListIds = projectedPlan.OrderedSourceListIds
                        .Where(selectedListIds.Contains)
                        .ToList();
                    projectedPlan.Lists = projectedPlan.Lists
                        .Where(value => selectedListIds.Contains(value.SourceListId))
                        .ToList();
                    var freshLists = ListMigrationTargetAnalyzer.InspectFresh(
                        targetContext,
                        projectedSources,
                        projectedPlan,
                        freshTopology);
                    blockers.AddRange(freshLists.Issues.Select(value => value.Code + ": " + value.Message));
                }
            }
            PublishingPageLayoutTargetProbe freshLayoutProbe = null;
            if (executionScope.Layout)
            {
                freshLayoutProbe = PublishingPageLayoutTargetInspector.Inspect(
                    targetContext,
                    package.Plan.LayoutMaterialization);
                var freshLayoutAdmission = PublishingPageLayoutTargetAdmissionEvaluator.Evaluate(
                    package.Plan.LayoutMaterialization,
                    freshLayoutProbe);
                foreach (var issue in freshLayoutAdmission.Issues)
                {
                    blockers.Add($"{issue.Code}: {issue.Message}");
                }
            }

            PublishingPageTargetSnapshot freshProbe = null;
            if (executionScope.PageArtifact)
            {
                var pageBlockers = new List<string>();
                freshProbe = PublishingPageTargetInspector.Inspect(
                    targetContext,
                    package.Plan.TargetPageServerRelativeUrl,
                    executionScope.ReferenceActions(package),
                    executionScope.Lifecycle
                        ? package.Plan.TargetLifecycle
                        : PnP.Framework.Migration.Pages.Publishing.Lifecycle.PublishingPageTargetLifecycle.Draft,
                    package.Plan.LayoutMaterialization,
                    freshLayoutProbe,
                    pageBlockers,
                    freshTopology);
                if (freshProbe.TargetPageExists
                    && IsOwnedByApprovedPlan(targetContext, package))
                {
                    pageBlockers.RemoveAll(value => value.StartsWith(
                        "Create-only target page already exists:",
                        StringComparison.Ordinal));
                }
                blockers.AddRange(pageBlockers);
            }
            if (blockers.Count > 0)
            {
                return Failure(package, operationId, startedAt, freshProbe == null ? "TargetPreflightFailed" : DetermineFailureCode(freshProbe), package.Plan.TargetPageServerRelativeUrl,
                    "Fresh target preflight failed: " + string.Join(" ", blockers), recorder);
            }

            if (freshProbe != null
                && (!PublishingPageContentTypeIdentity.MatchesSiteContentType(
                        freshProbe.PageContentTypeId,
                        package.Plan.TargetProbe.PageContentTypeId)
                || !string.Equals(freshProbe.PageLayoutUrl, package.Plan.TargetProbe.PageLayoutUrl, StringComparison.OrdinalIgnoreCase))
                )
            {
                return Failure(package, operationId, startedAt, "TargetPreconditionChanged", package.Plan.TargetPageServerRelativeUrl,
                    "The target Publishing Page content type or layout changed after approval.", recorder);
            }

            return null;
        }

        private static bool IsOwnedByApprovedPlan(
            ClientContext context,
            PublishingPageMigrationPackage package)
        {
            var file = context.Web.GetFileByServerRelativePath(
                ResourcePath.FromDecodedUrl(package.Plan.TargetPageServerRelativeUrl));
            context.Load(file, value => value.Exists, value => value.Properties);
            try
            {
                context.ExecuteQueryRetry();
            }
            catch (ServerException exception) when (
                string.Equals(exception.ServerErrorTypeName, "System.IO.FileNotFoundException", StringComparison.Ordinal)
                || exception.ServerErrorCode == -2147024894)
            {
                return false;
            }

            return file.Exists
                && PublishingPageTargetOwnership.MatchesApprovedPlan(
                    file.Properties.FieldValues,
                    package.Plan.OriginalIdentifier,
                    package.SnapshotDigest,
                    package.PlanDigest);
        }

        private static bool SameAuthority(string left, string right)
        {
            return Uri.TryCreate(left, UriKind.Absolute, out var leftUri)
                && Uri.TryCreate(right, UriKind.Absolute, out var rightUri)
                && string.Equals(leftUri.Scheme, rightUri.Scheme, StringComparison.OrdinalIgnoreCase)
                && string.Equals(leftUri.Authority, rightUri.Authority, StringComparison.OrdinalIgnoreCase);
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
