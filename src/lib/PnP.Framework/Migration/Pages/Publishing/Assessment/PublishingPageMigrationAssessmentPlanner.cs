using PnP.Framework.Migration.Lists.Planning;
using PnP.Framework.Migration.Packaging;
using PnP.Framework.Migration.Pages.Assessment;
using PnP.Framework.Migration.Pages.Cohorts;
using PnP.Framework.Migration.Pages.Planning;
using PnP.Framework.Migration.Pages.Publishing.Ingredients;
using PnP.Framework.Migration.Pages.Publishing.Layouts;
using PnP.Framework.Migration.Pages.Publishing.Lifecycle;
using PnP.Framework.Migration.Pages.Publishing.Packaging;
using PnP.Framework.Migration.Pages.Publishing.Planning;
using PnP.Framework.Migration.Pages.Publishing.Profiles;
using PnP.Framework.Migration.Pages.References;
using PnP.Framework.Migration.Topology;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace PnP.Framework.Migration.Pages.Publishing.Assessment
{
    internal sealed class PublishingPageMigrationAssessmentPlanner
    {
        public PublishingPageMigrationAssessment Assess(
            PublishingPageExportPackage exportPackage,
            TopologyPlan topology,
            PagePlanningOptions options,
            PublishingPageWorkflowPolicy workflowPolicy,
            IMigrationArtifactStore artifactStore,
            PageAssessmentEvidence assessmentEvidence = null)
        {
            if (workflowPolicy == null)
            {
                throw new ArgumentNullException(nameof(workflowPolicy));
            }

            PublishingPagePackageValidator.ValidateExport(exportPackage, artifactStore);
            TopologyPlanValidator.Validate(topology);
            PublishingPagePlanningPolicy.ValidateOptions(options);
            ValidateWorkflow(exportPackage, workflowPolicy);

            var assessmentOptions = PublishingPagePlanningPolicy.CopyOptions(
                options,
                options.TargetPageServerRelativeUrl);
            var taxonomyReviewPlan = assessmentEvidence?.TaxonomyAssetReviewPlan;
            if (taxonomyReviewPlan != null)
            {
                assessmentOptions.TaxonomySchemaMappings =
                    PagePlanningTaxonomyMappingResolver.ResolveForAssessment(
                        assessmentOptions.TaxonomySchemaMappings,
                        taxonomyReviewPlan);
                // Prospective assessment mappings must never masquerade as the
                // post-materialization, fresh-readback mapping catalog.
                assessmentOptions.TaxonomyAssetMappingCatalog = null;
            }

            var snapshot = exportPackage.Snapshot;
            var targetSite = topology.SiteCollections.SingleOrDefault(value =>
                value.SourceSiteId == snapshot.Source.SiteId);
            var targetWeb = targetSite?.Webs.SingleOrDefault(value =>
                value.SourceWebId == snapshot.Source.WebId);
            var targetPagePath = targetWeb == null
                ? options.TargetPageServerRelativeUrl
                : PagePath.Normalize(
                    targetWeb.TargetServerRelativeUrl,
                    options.TargetPageServerRelativeUrl,
                    "Pages");
            var sourceCaptureDiagnostics = (snapshot.Blockers ?? Array.Empty<string>()).
                Where(value => !string.IsNullOrWhiteSpace(value)).ToList();
            var knownGaps = new List<string>();
            var warnings = (snapshot.Warnings ?? Array.Empty<string>()).
                Where(value => !string.IsNullOrWhiteSpace(value)).ToList();
            if (taxonomyReviewPlan != null)
            {
                warnings.Add(
                    "Taxonomy assessment uses deterministic candidates from read-only review plan '"
                    + taxonomyReviewPlan.PlanDigest
                    + "'. This evidence does not authorize taxonomy or page mutation; execution still requires approval, materialization, and fresh-readback admission.");
            }
            warnings.AddRange(sourceCaptureDiagnostics.Select(value =>
                "Source capture diagnostic (classified by the ingredient assessment): " + value));
            if (exportPackage.Selection.ValidationCohort?.Disposition != ValidationCohortDisposition.Included)
            {
                warnings.Add(
                    $"The source page is '{exportPackage.Selection.ValidationCohort?.Disposition}' for validation cohort "
                    + $"'{exportPackage.Selection.ValidationCohort?.CohortId}'; CLR runtime and ingredient capability remain authoritative.");
            }

            var targetLifecycle = PublishingPageLifecyclePolicy.DeriveTargetLifecycle(snapshot.Lifecycle);
            var lifecycleReason = PublishingPagePlanningPolicy.DescribeLifecycleDecision(
                snapshot.Lifecycle,
                targetLifecycle,
                warnings);
            var context = new PublishingPageAssessmentContext
            {
                Snapshot = snapshot,
                WorkflowPolicy = workflowPolicy,
                Options = assessmentOptions,
                TaxonomyAssetReviewPlan = taxonomyReviewPlan,
                TargetSite = targetSite,
                TargetWeb = targetWeb,
                TargetPageServerRelativeUrl = targetPagePath,
                TargetLifecycle = targetLifecycle,
                LifecycleReason = lifecycleReason,
                KnownGaps = knownGaps,
                Warnings = warnings
            };

            BuildLayoutPlan(context, artifactStore);
            BuildReferenceActions(context);
            BuildListPlan(context, topology);

            var graph = PublishingPageIngredientGraphProjector.Project(snapshot);
            var accumulator = new PublishingPageAssessmentAccumulator(graph);
            PublishingPageCoreAssessmentProjector.Project(context, accumulator);
            PublishingPageLayoutAssessmentProjector.Project(context, accumulator);
            PublishingPageTopologyAssessmentProjector.Project(context, accumulator);
            PublishingPageListAssessmentProjector.Project(context, accumulator);
            PublishingPageWebPartAssessmentProjector.Project(context, accumulator);
            PublishingPageReferenceAssessmentProjector.Project(context, accumulator);
            var ingredientAssessments = accumulator.Complete();
            PublishingPageAuthorizationEvidenceProjector.Apply(
                ingredientAssessments,
                PublishingPageSnapshotAuthorizationEvidence.Merge(snapshot, assessmentEvidence));
            knownGaps.AddRange(ingredientAssessments
                .Where(value => value.State == PageIngredientAssessmentState.KnownGap)
                .Select(value => value.MitigationCode + ": " + value.Reason));

            var assessment = new PublishingPageMigrationAssessment
            {
                SourceSnapshotDigest = exportPackage.SnapshotDigest,
                WorkflowId = exportPackage.Selection.WorkflowId,
                SelectionDigest = exportPackage.SelectionDigest,
                SourceWebUrl = snapshot.Source.WebUrl,
                SourcePageServerRelativeUrl = snapshot.Source.PageServerRelativeUrl,
                TargetSiteCollectionUrl = targetSite?.TargetSiteCollectionUrl,
                TargetWebUrl = targetWeb?.TargetWebUrl,
                TargetWebServerRelativeUrl = targetWeb?.TargetServerRelativeUrl,
                TargetPageServerRelativeUrl = targetPagePath,
                TopologyPlanDigest = topology.PlanDigest,
                PlanningPolicy = PublishingPagePlanningPolicy.CopyOptions(assessmentOptions, targetPagePath),
                TargetLifecycle = targetLifecycle,
                LifecycleReason = lifecycleReason,
                IngredientGraph = graph,
                IngredientAssessments = ingredientAssessments,
                KnownGaps = knownGaps.Distinct(StringComparer.Ordinal).OrderBy(value => value, StringComparer.Ordinal).ToList(),
                Warnings = warnings.Distinct(StringComparer.Ordinal).OrderBy(value => value, StringComparer.Ordinal).ToList()
            };
            assessment.State = assessment.IngredientAssessments.Any(value =>
                    value.State == PageIngredientAssessmentState.AuthorizationBlocked)
                ? PageMigrationAssessmentState.AuthorizationBlocked
                : assessment.IngredientAssessments.Any(value =>
                        value.State == PageIngredientAssessmentState.KnownGap)
                    || assessment.KnownGaps.Count > 0
                        ? PageMigrationAssessmentState.KnownGap
                        : PageMigrationAssessmentState.ReadyForTargetInspection;
            assessment.AssessmentDigest = PublishingPageAssessmentDigest.Compute(assessment);
            PublishingPageMigrationAssessmentValidator.Validate(assessment);
            return assessment;
        }

        private static void ValidateWorkflow(
            PublishingPageExportPackage exportPackage,
            PublishingPageWorkflowPolicy workflowPolicy)
        {
            if (!string.Equals(exportPackage.Selection.WorkflowId, workflowPolicy.WorkflowId, StringComparison.Ordinal))
            {
                throw new InvalidDataException(
                    $"Workflow '{exportPackage.Selection.WorkflowId}' cannot be assessed by policy '{workflowPolicy.WorkflowId}'.");
            }
            var expected = workflowPolicy.Select(exportPackage.Snapshot.Source.ContentTypeId);
            if (!string.Equals(
                    PublishingPageDigest.ComputeSelectionDigest(expected),
                    exportPackage.SelectionDigest,
                    StringComparison.OrdinalIgnoreCase))
            {
                throw new InvalidDataException(
                    "The sealed validation-cohort assessment does not match the selected workflow policy and source evidence.");
            }
        }

        private static void BuildLayoutPlan(
            PublishingPageAssessmentContext context,
            IMigrationArtifactStore artifactStore)
        {
            if (context.TargetSite == null || context.TargetWeb == null)
            {
                context.LayoutPlanningFailure =
                    "The reviewed topology plan has no target Site/Web mapping for the source page.";
                context.KnownGaps.Add("TargetWebTopologyMappingUnavailable: " + context.LayoutPlanningFailure);
                return;
            }
            try
            {
                context.LayoutPlan = PublishingPageLayoutPlanFactory.Create(
                    context.Snapshot.Layout,
                    new Uri(context.Snapshot.Source.WebUrl),
                    new Uri(context.TargetWeb.TargetWebUrl),
                    new Uri(context.TargetSite.TargetSiteCollectionUrl),
                    context.WorkflowPolicy.PreferredTargetPageLayoutFileName,
                    context.Options.TaxonomySchemaMappings,
                    artifactStore,
                    context.Options.AllowExternalResourceReferences);
            }
            catch (Exception exception)
            {
                context.LayoutPlanningFailure = exception.GetType().Name + ": " + exception.Message;
                context.KnownGaps.Add("PageLayoutPlanningFailed: " + context.LayoutPlanningFailure);
            }
        }

        private static void BuildReferenceActions(PublishingPageAssessmentContext context)
        {
            if (context.TargetWeb == null)
            {
                context.ReferencePlanningFailure =
                    "The reviewed topology plan has no target Web mapping for source references.";
                if (context.Snapshot.Dependencies.Count > 0)
                {
                    context.KnownGaps.Add("PageReferencePlanningFailed: " + context.ReferencePlanningFailure);
                }
                return;
            }
            try
            {
                context.ReferenceActions = PageReferencePlanner.BuildActions(
                    context.Snapshot.Source,
                    context.Snapshot.Dependencies,
                    context.TargetWeb.TargetWebUrl,
                    context.TargetWeb.TargetServerRelativeUrl,
                    context.TargetSite,
                    context.Options,
                    context.KnownGaps);
                context.Replacements = PageReferencePlanner.BuildTextReplacements(
                    context.Snapshot.Source,
                    context.TargetWeb.TargetWebUrl,
                    context.TargetWeb.TargetServerRelativeUrl,
                    context.Snapshot.Dependencies,
                    context.ReferenceActions);
            }
            catch (Exception exception)
            {
                context.ReferencePlanningFailure = exception.GetType().Name + ": " + exception.Message;
                context.KnownGaps.Add("PageReferencePlanningFailed: " + context.ReferencePlanningFailure);
            }
        }

        private static void BuildListPlan(
            PublishingPageAssessmentContext context,
            TopologyPlan topology)
        {
            try
            {
                context.ListPlan = ListMigrationPlanFactory.Create(
                    context.Snapshot.ListDependencies,
                    context.Snapshot.ListLookupDependencies,
                    topology,
                    context.Options.TaxonomySchemaMappings,
                    context.Options.ListTargetOverrides);
            }
            catch (Exception exception)
            {
                context.ListPlanningFailure = exception.GetType().Name + ": " + exception.Message;
                if (context.Snapshot.ListDependencies.Count > 0)
                {
                    context.KnownGaps.Add("ListPlanningFailed: " + context.ListPlanningFailure);
                }
            }
        }
    }
}
