using Microsoft.SharePoint.Client;
using PnP.Framework.Migration.Pages.Capture;
using PnP.Framework.Migration.Pages.Content;
using PnP.Framework.Migration.Pages.Fields;
using PnP.Framework.Migration.Pages.Lifecycle;
using PnP.Framework.Migration.Pages.Publishing.Packaging;
using PnP.Framework.Migration.Pages.Planning;
using PnP.Framework.Migration.Pages.References;
using PnP.Framework.Migration.Pages.Publishing.Capture;
using PnP.Framework.Migration.Pages.Publishing.Lifecycle;
using PnP.Framework.Migration.Pages.Publishing.Layouts;
using PnP.Framework.Migration.Pages.Publishing.Planning;
using PnP.Framework.Migration.Pages.Publishing.Reporting;
using PnP.Framework.Migration.Pages.Publishing.Verification;
using PnP.Framework.Migration.Verification;
using PnP.Framework.Migration.Packaging;
using PnP.Framework.Migration.Taxonomy;
using PnP.Framework.Migration.Topology;
using PnP.Framework.Migration.Lists.Planning;
using PnP.Framework.Migration.Pages.ClassicWebParts.Bindings;
using PnP.Framework.Migration.Pages.ClassicWebParts.Planning;
using PnP.Framework.Migration.Pages.Cohorts;
using PnP.Framework.Migration.Pages.Ingredients;
using PnP.Framework.Migration.Pages.Publishing.Ingredients;
using PnP.Framework.Migration.Pages.Publishing.Profiles;
using PnP.Framework.Migration.Pages.Runtime;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;

namespace PnP.Framework.Migration.Pages.Publishing.Planning
{
    internal sealed class PublishingPageMigrationPlanner
    {
        public PublishingPageMigrationPackage Plan(
            ClientContext targetContext,
            PublishingPageExportPackage exportPackage,
            PagePlanningOptions options,
            PublishingPageWorkflowPolicy workflowPolicy)
        {
            return Plan(targetContext, exportPackage, options, workflowPolicy, null);
        }

        public PublishingPageMigrationPackage Plan(
            ClientContext targetContext,
            PublishingPageExportPackage exportPackage,
            PagePlanningOptions options,
            PublishingPageWorkflowPolicy workflowPolicy,
            IMigrationArtifactStore artifactStore)
        {
            if (targetContext == null)
            {
                throw new ArgumentNullException(nameof(targetContext));
            }
            if (workflowPolicy == null)
            {
                throw new ArgumentNullException(nameof(workflowPolicy));
            }

            PublishingPagePackageValidator.ValidateExport(exportPackage, artifactStore);
            if (!string.Equals(exportPackage.Selection.WorkflowId, workflowPolicy.WorkflowId, StringComparison.Ordinal))
            {
                throw new InvalidDataException($"Workflow '{exportPackage.Selection.WorkflowId}' cannot be planned by policy '{workflowPolicy.WorkflowId}'.");
            }
            var expectedSelection = workflowPolicy.Select(exportPackage.Snapshot.Source.ContentTypeId);
            if (!string.Equals(
                    PublishingPageDigest.ComputeSelectionDigest(expectedSelection),
                    exportPackage.SelectionDigest,
                    StringComparison.OrdinalIgnoreCase))
            {
                throw new InvalidDataException("The sealed validation-cohort assessment does not match the selected workflow policy and source evidence.");
            }
            PublishingPagePlanningPolicy.ValidateOptions(options);
            var targetWeb = targetContext.Web;
            var targetSite = targetContext.Site;
            var targetRootWeb = targetContext.Site.RootWeb;
            targetContext.Load(targetWeb,
                web => web.Id,
                web => web.Url,
                web => web.ServerRelativeUrl,
                web => web.Title,
                web => web.WebTemplate,
                web => web.Configuration,
                web => web.AllProperties);
            targetContext.Load(targetSite, site => site.Id, site => site.ServerRelativeUrl);
            targetContext.Load(targetRootWeb,
                web => web.Id,
                web => web.Url,
                web => web.ServerRelativeUrl,
                web => web.Title,
                web => web.WebTemplate,
                web => web.Configuration);
            targetContext.ExecuteQueryRetry();

            var snapshot = exportPackage.Snapshot;
            var targetPagePath = PagePath.Normalize(targetWeb.ServerRelativeUrl, options.TargetPageServerRelativeUrl, "Pages");
            var pageOriginalIdentifier = PublishingPageTargetOwnership.OriginalIdentifier(snapshot.Source);
            var blockers = snapshot.Blockers.ToList();
            var warnings = snapshot.Warnings.ToList();
            if (exportPackage.Selection?.ValidationCohort?.Disposition != ValidationCohortDisposition.Included)
            {
                warnings.Add($"The source page is '{exportPackage.Selection.ValidationCohort.Disposition}' for validation cohort '{exportPackage.Selection.ValidationCohort.CohortId}'. Cohort membership is retained as evidence but does not override the CLR-selected Publishing runtime or ingredient capability decisions.");
            }
            if (!string.Equals(snapshot.Runtime?.AdapterId, PageRuntimeAdapterIds.Publishing, StringComparison.Ordinal))
            {
                blockers.Add($"The detected runtime adapter '{snapshot.Runtime?.AdapterId ?? PageRuntimeAdapterIds.Unknown}' is not executable by the Publishing Page planner.");
            }
            PublishingPagePlanningPolicy.AddSnapshotDecisions(snapshot, options, blockers);

            var layoutMaterialization = PublishingPageLayoutPlanFactory.Create(
                snapshot.Layout,
                new Uri(snapshot.Source.WebUrl),
                new Uri(targetWeb.Url),
                new Uri(targetRootWeb.Url),
                workflowPolicy.PreferredTargetPageLayoutFileName,
                options.TaxonomySchemaMappings,
                artifactStore,
                options.AllowExternalResourceReferences);
            var layoutTargetProbe = layoutMaterialization.Disposition == PublishingPageLayoutMaterializationDisposition.Block
                ? null
                : PublishingPageLayoutTargetInspector.Inspect(targetContext, layoutMaterialization);
            var layoutAdmission = PublishingPageLayoutTargetAdmissionEvaluator.Evaluate(layoutMaterialization, layoutTargetProbe);
            foreach (var issue in layoutAdmission.Issues)
            {
                blockers.Add($"{issue.Code}: {issue.Message}");
            }
            foreach (var warning in layoutAdmission.Warnings)
            {
                warnings.Add(warning);
            }

            var targetLifecycle = PublishingPageLifecyclePolicy.DeriveTargetLifecycle(snapshot.Lifecycle);
            var lifecycleReason = PublishingPagePlanningPolicy.DescribeLifecycleDecision(snapshot.Lifecycle, targetLifecycle, warnings);
            var dependencyPlan = PublishingPageDependencyPlanner.Build(
                targetContext,
                snapshot,
                targetWeb,
                targetSite,
                targetRootWeb,
                options,
                blockers,
                warnings);
            var referenceSiteMapping = dependencyPlan.Topology?.SiteCollections
                .SingleOrDefault(value => value.SourceSiteId == snapshot.Source.SiteId);
            var dependencyActions = PageReferencePlanner.BuildActions(
                snapshot.Source,
                snapshot.Dependencies,
                targetWeb.Url,
                targetWeb.ServerRelativeUrl,
                referenceSiteMapping,
                options,
                blockers);
            var replacements = PageReferencePlanner.BuildTextReplacements(
                snapshot.Source,
                targetWeb.Url,
                targetWeb.ServerRelativeUrl,
                snapshot.Dependencies,
                dependencyActions);
            Microsoft.SharePoint.Client.List targetPages;
            var targetProbe = PublishingPageTargetInspector.InspectForPlanning(
                targetContext,
                targetPagePath,
                pageOriginalIdentifier,
                dependencyActions,
                targetLifecycle,
                layoutMaterialization,
                layoutTargetProbe,
                blockers,
                out targetPages,
                includeListInventory: snapshot.ListDependencies.Count > 0,
                dependencyTopology: dependencyPlan.TopologyTargetAnalysis);
            if (targetProbe.TargetPathCollisionResolved)
            {
                targetPagePath = targetProbe.TargetPageServerRelativeUrl;
                warnings.Add(targetProbe.TargetPathResolutionReason + " Final target: '" + targetPagePath + "'.");
            }
            var taxonomyRelationshipActions = new List<TaxonomyRelationshipAction>();
            var fieldActions = PageFieldPlanner.BuildActions(
                targetContext,
                snapshot.Fields,
                workflowPolicy.FieldsHandledByPageWriter,
                workflowPolicy.RecognizedPageFields,
                options,
                taxonomyRelationshipActions,
                blockers,
                warnings,
                targetPages,
                targetPagesResolved: true,
                targetFieldsLoaded: targetPages != null);
            var expectedContent = PageTextTransformer.Rewrite(snapshot.PublishingPageContent, replacements);
            var expectedContentDigest = PublishingPageDigest.ComputeSha256(expectedContent);
            var planningIngredientGraph = PublishingPageIngredientGraphProjector.Project(snapshot);
            var plan = new PublishingPageMigrationPlan
            {
                SourceSnapshotDigest = exportPackage.SnapshotDigest,
                SourceWebUrl = snapshot.Source.WebUrl,
                SourcePageServerRelativeUrl = snapshot.Source.PageServerRelativeUrl,
                OriginalIdentifier = pageOriginalIdentifier,
                TargetWebUrl = targetWeb.Url.TrimEnd('/'),
                TargetWebServerRelativeUrl = targetWeb.ServerRelativeUrl,
                PreferredTargetPageServerRelativeUrl = targetProbe.PreferredTargetPageServerRelativeUrl,
                TargetPageServerRelativeUrl = targetPagePath,
                TargetPathCollisionResolved = targetProbe.TargetPathCollisionResolved,
                TargetPathResolutionReason = targetProbe.TargetPathResolutionReason,
                PageLayoutName = layoutMaterialization.TargetPageLayoutName,
                Operation = PageMigrationOperation.CreatePage,
                TargetLifecycle = targetLifecycle,
                LifecycleReason = lifecycleReason,
                CreateOnly = options.CreateOnly,
                PlanningPolicy = PublishingPagePlanningPolicy.CopyOptions(options, targetPagePath),
                TargetProbe = targetProbe,
                LayoutMaterialization = layoutMaterialization,
                LayoutTargetProbe = layoutTargetProbe,
                LayoutAdmission = layoutAdmission,
                FieldActions = fieldActions,
                TaxonomyRelationshipActions = taxonomyRelationshipActions,
                DependencyActions = dependencyActions,
                Topology = dependencyPlan.Topology,
                TopologyTargetAnalysis = dependencyPlan.TopologyTargetAnalysis,
                ListMigration = dependencyPlan.ListMigration,
                WebPartActions = dependencyPlan.WebPartActions,
                Replacements = replacements,
                ExpectedPublishingPageContentSha256 = expectedContentDigest,
                StorageAssertions = PageStorageAssertionBuilder.Build(
                    snapshot,
                    targetPagePath,
                    dependencyActions,
                    expectedContentDigest,
                    targetLifecycle),
                RuntimeVerification = PublishingPageRuntimeVerificationPolicy.CreateManifest(),
                IngredientGraph = planningIngredientGraph,
                Blockers = blockers.Distinct(StringComparer.Ordinal).OrderBy(item => item, StringComparer.Ordinal).ToList(),
                Warnings = warnings.Distinct(StringComparer.Ordinal).OrderBy(item => item, StringComparer.Ordinal).ToList()
            };
            plan.IngredientActions = PublishingPageIngredientActionProjector.Project(snapshot, plan, planningIngredientGraph);
            var ingredientEvaluation = PageIngredientPlanEvaluator.Evaluate(
                planningIngredientGraph,
                plan.IngredientActions,
                PublishingPageIngredientAuthorizationPolicy.GetEvidence(snapshot));
            plan.MigrationOutcome = ingredientEvaluation.Outcome;
            plan.IngredientIssues = ingredientEvaluation.Issues;
            plan.ExecutionFrontier = ingredientEvaluation.ExecutionFrontier;
            var package = new PublishingPageMigrationPackage
            {
                PlannedAtUtc = DateTimeOffset.UtcNow,
                ExportedAtUtc = exportPackage.ExportedAtUtc,
                Selection = exportPackage.Selection,
                SelectionDigest = exportPackage.SelectionDigest,
                State = PublishingPagePackageStatePolicy.Derive(plan),
                Snapshot = snapshot,
                Plan = plan,
                SnapshotDigest = exportPackage.SnapshotDigest,
                PlanDigest = PublishingPageDigest.ComputePlanDigest(plan),
                Report = PublishingPagePlanReportFactory.Create(snapshot, plan)
            };
            PublishingPagePackageValidator.ValidateMigration(package, artifactStore);
            return package;
        }

    }
}
