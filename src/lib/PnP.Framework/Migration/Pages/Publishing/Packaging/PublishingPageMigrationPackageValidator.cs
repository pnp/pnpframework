using PnP.Framework.Migration.Diagnostics;
using PnP.Framework.Migration.Lists.Packaging;
using PnP.Framework.Migration.Packaging;
using PnP.Framework.Migration.Pages.Content;
using PnP.Framework.Migration.Pages.Ingredients;
using PnP.Framework.Migration.Pages.Publishing.Capture;
using PnP.Framework.Migration.Pages.Publishing.Ingredients;
using PnP.Framework.Migration.Pages.Publishing.Layouts.Packaging;
using PnP.Framework.Migration.Pages.Publishing.Lifecycle;
using PnP.Framework.Migration.Pages.Publishing.Planning;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace PnP.Framework.Migration.Pages.Publishing.Packaging
{
    internal static class PublishingPageMigrationPackageValidator
    {
        public static void Validate(
            PublishingPageMigrationPackage package,
            IMigrationArtifactStore artifactStore)
        {
            ValidateEnvelope(package);
            PublishingPageExportPackageValidator.Validate(new PublishingPageExportPackage
            {
                SchemaVersion = package.ExportSchemaVersion,
                ExportedAtUtc = package.ExportedAtUtc,
                Selection = package.Selection,
                SelectionDigest = package.SelectionDigest,
                Snapshot = package.Snapshot,
                SnapshotDigest = package.SnapshotDigest
            }, artifactStore);

            var plan = package.Plan;
            ValidatePlanShape(plan);
            PublishingPageLayoutPackageValidator.ValidatePlan(
                plan.PageLayoutName,
                plan.IsExecutable,
                plan.LayoutMaterialization,
                plan.LayoutTargetProbe,
                plan.LayoutAdmission);
            if (plan.IsExecutable && string.IsNullOrWhiteSpace(plan.TargetProbe.PageContentTypeId))
            {
                throw new InvalidDataException("An executable Publishing Page plan must seal one exact target Pages-library Content Type ID.");
            }

            ValidateTopology(package, plan);
            ListMigrationPlanValidator.Validate(package.Snapshot.ListDependencies, plan.ListMigration);
            ValidateRuntimeVerification(plan);
            if (!string.Equals(plan.SourceSnapshotDigest, package.SnapshotDigest, StringComparison.OrdinalIgnoreCase))
            {
                throw new InvalidDataException("The migration plan does not reference the sealed snapshot in this package.");
            }
            var planDigest = PublishingPageDigest.ComputePlanDigest(plan);
            if (!string.Equals(planDigest, package.PlanDigest, StringComparison.OrdinalIgnoreCase))
            {
                throw new InvalidDataException("The migration plan digest does not match the package payload.");
            }

            ValidateActionCoverage(package.Snapshot, plan);
            ValidateDerivedIngredientActions(package.Snapshot, plan);
            ValidateIngredientActions(package.Snapshot.IngredientGraph, plan);
            ValidateExpectedContent(package, plan);
            var derivedLifecycle = PublishingPageLifecyclePolicy.DeriveTargetLifecycle(package.Snapshot.Lifecycle);
            if (plan.TargetLifecycle != derivedLifecycle)
            {
                throw new InvalidDataException($"The planned lifecycle '{plan.TargetLifecycle}' does not match the source-derived lifecycle '{derivedLifecycle}'.");
            }
            var expectedState = plan.IsExecutable
                ? PublishingPagePackageState.ApprovalReady
                : PublishingPagePackageState.Blocked;
            if (package.State != expectedState)
            {
                throw new InvalidDataException($"Package state '{package.State}' does not match plan executability '{expectedState}'.");
            }
        }

        private static void ValidateEnvelope(PublishingPageMigrationPackage package)
        {
            if (package == null)
            {
                throw new InvalidDataException("The publishing-page migration package is empty.");
            }
            if (!string.Equals(package.SchemaVersion, PublishingPagePackageContract.MigrationSchemaVersion, StringComparison.Ordinal))
            {
                throw new InvalidDataException($"Unsupported publishing-page migration schema '{package.SchemaVersion}'.");
            }
            if (!string.Equals(package.ExportSchemaVersion, PublishingPagePackageContract.ExportSchemaVersion, StringComparison.Ordinal))
            {
                throw new InvalidDataException($"Unsupported embedded publishing-page export schema '{package.ExportSchemaVersion}'.");
            }
            if (package.Snapshot == null || package.Plan == null)
            {
                throw new InvalidDataException("The migration package must contain both a snapshot and a plan.");
            }
        }

        private static void ValidatePlanShape(PublishingPageMigrationPlan plan)
        {
            if (plan.PlanningPolicy == null
                || plan.TargetProbe == null
                || plan.LayoutMaterialization == null
                || plan.LayoutAdmission == null
                || plan.FieldActions == null
                || plan.DependencyActions == null
                || plan.WebPartActions == null
                || plan.Replacements == null
                || plan.StorageAssertions == null
                || plan.RuntimeVerification == null
                || plan.RuntimeVerification.Requirements == null
                || plan.IngredientActions == null
                || plan.IngredientIssues == null
                || plan.Blockers == null
                || plan.Warnings == null)
            {
                throw new InvalidDataException("The migration plan is missing policy, target probe, or an action/assertion collection.");
            }
            if (plan.PlanningPolicy.TaxonomySchemaMappings == null
                || plan.PlanningPolicy.TopologyPolicy == null
                || plan.PlanningPolicy.TopologyPolicy.WebOverrides == null
                || plan.PlanningPolicy.ListTargetOverrides == null)
            {
                throw new InvalidDataException("The planning policy contains a null taxonomy schema mapping collection.");
            }
        }

        private static void ValidateTopology(
            PublishingPageMigrationPackage package,
            PublishingPageMigrationPlan plan)
        {
            if (package.Snapshot.SourceTopology != null && plan.Topology != null
                && !string.Equals(Topology.TopologyPlanner.ComputeDigest(plan.Topology), plan.Topology.PlanDigest, StringComparison.OrdinalIgnoreCase))
            {
                throw new InvalidDataException("The topology plan digest differs from its sealed content.");
            }
            if (plan.Topology == null)
            {
                return;
            }
            if (plan.TopologyTargetAnalysis == null
                || !string.Equals(plan.TopologyTargetAnalysis.TopologyPlanDigest, plan.Topology.PlanDigest, StringComparison.OrdinalIgnoreCase))
            {
                throw new InvalidDataException("The topology plan requires target analysis sealed against the same topology digest.");
            }
            var plannedWebs = new HashSet<Guid>(plan.Topology.SiteCollections.SelectMany(value => value.Webs).Select(value => value.SourceWebId));
            var probedWebs = new HashSet<Guid>(plan.TopologyTargetAnalysis.SiteCollections.SelectMany(value => value.Webs).Select(value => value.SourceWebId));
            var plannedWebCount = plan.Topology.SiteCollections.SelectMany(value => value.Webs).Count();
            var probedWebCount = plan.TopologyTargetAnalysis.SiteCollections.SelectMany(value => value.Webs).Count();
            var plannedSiteIds = plan.Topology.SiteCollections.Select(value => value.SourceSiteId).ToArray();
            var probedSiteIds = plan.TopologyTargetAnalysis.SiteCollections.Select(value => value.SourceSiteId).ToArray();
            if (plannedWebCount != plannedWebs.Count
                || probedWebCount != probedWebs.Count
                || plannedSiteIds.Length != plannedSiteIds.Distinct().Count()
                || probedSiteIds.Length != probedSiteIds.Distinct().Count()
                || !plannedWebs.SetEquals(probedWebs)
                || !new HashSet<Guid>(plannedSiteIds).SetEquals(probedSiteIds))
            {
                throw new InvalidDataException("The target topology analysis must cover every planned Web exactly once.");
            }
        }

        private static void ValidateRuntimeVerification(PublishingPageMigrationPlan plan)
        {
            var duplicate = plan.RuntimeVerification.Requirements
                .GroupBy(item => item?.Id, StringComparer.Ordinal)
                .FirstOrDefault(group => string.IsNullOrWhiteSpace(group.Key) || group.Count() > 1);
            if (duplicate != null || plan.RuntimeVerification.Requirements.Any(item => item == null))
            {
                throw new InvalidDataException($"The runtime verification manifest contains a missing or duplicate requirement ID '{duplicate?.Key}'.");
            }
        }

        private static void ValidateExpectedContent(
            PublishingPageMigrationPackage package,
            PublishingPageMigrationPlan plan)
        {
            var expectedContent = PageTextTransformer.Rewrite(
                package.Snapshot.PublishingPageContent,
                plan.Replacements);
            if (!string.Equals(
                    PublishingPageDigest.ComputeSha256(expectedContent),
                    plan.ExpectedPublishingPageContentSha256,
                    StringComparison.OrdinalIgnoreCase))
            {
                throw new InvalidDataException("The expected target PublishingPageContent digest does not match the approved replacements.");
            }
        }

        private static void ValidateDerivedIngredientActions(
            PublishingPageCaptureBundle snapshot,
            PublishingPageMigrationPlan plan)
        {
            var expected = PublishingPageIngredientActionProjector.Project(snapshot, plan);
            if (!PublishingPageValidationCanonical.Equals(expected, plan.IngredientActions))
            {
                throw new InvalidDataException("The sealed ingredient actions do not match the typed domain plans and policy projection.");
            }
        }

        private static void ValidateIngredientActions(
            CanonicalPageIngredientGraph graph,
            PublishingPageMigrationPlan plan)
        {
            if (plan.IngredientActions.Any(action => action == null
                    || string.IsNullOrWhiteSpace(action.ActionId)
                    || string.IsNullOrWhiteSpace(action.PolicyId)
                    || string.IsNullOrWhiteSpace(action.PolicyVersion)
                    || action.ReleasedDependencyIngredientIds == null
                    || action.VerificationAssertions == null))
            {
                throw new InvalidDataException("Every ingredient action must have an ID and non-null dependency/assertion collections.");
            }
            var duplicateActionId = plan.IngredientActions
                .GroupBy(action => action.ActionId, StringComparer.Ordinal)
                .FirstOrDefault(group => group.Count() > 1);
            if (duplicateActionId != null)
            {
                throw new InvalidDataException($"Ingredient action ID '{duplicateActionId.Key}' is duplicated.");
            }
            var evaluation = PageIngredientPlanEvaluator.Evaluate(graph, plan.IngredientActions);
            if (evaluation.Outcome != plan.MigrationOutcome)
            {
                throw new InvalidDataException($"The sealed migration outcome '{plan.MigrationOutcome}' differs from evaluated outcome '{evaluation.Outcome}'.");
            }
            var expectedIssues = new HashSet<string>(evaluation.Issues.Select(IssueIdentity), StringComparer.Ordinal);
            var actualIssues = new HashSet<string>(plan.IngredientIssues.Select(IssueIdentity), StringComparer.Ordinal);
            if (plan.IngredientIssues.Any(value => value == null)
                || evaluation.Issues.Count != plan.IngredientIssues.Count
                || !expectedIssues.SetEquals(actualIssues))
            {
                throw new InvalidDataException("The sealed ingredient issues differ from dependency-closure evaluation.");
            }
        }

        private static void ValidateActionCoverage(
            PublishingPageCaptureBundle snapshot,
            PublishingPageMigrationPlan plan)
        {
            var sourceFieldNames = new HashSet<string>(snapshot.Fields.Select(item => item.InternalName), StringComparer.OrdinalIgnoreCase);
            var plannedFieldNames = new HashSet<string>(plan.FieldActions.Select(item => item?.SourceInternalName), StringComparer.OrdinalIgnoreCase);
            if (plan.FieldActions.Any(item => item == null)
                || plan.FieldActions.Count != sourceFieldNames.Count
                || plannedFieldNames.Count != sourceFieldNames.Count
                || !sourceFieldNames.SetEquals(plannedFieldNames))
            {
                throw new InvalidDataException("The plan must contain exactly one field action for every captured source field.");
            }
            var dependencyIds = new HashSet<string>(snapshot.Dependencies.Select(item => item.Id), StringComparer.Ordinal);
            var plannedDependencyIds = new HashSet<string>(plan.DependencyActions.Select(item => item?.SnapshotDependencyId), StringComparer.Ordinal);
            if (plan.DependencyActions.Any(item => item == null)
                || plan.DependencyActions.Count != dependencyIds.Count
                || plannedDependencyIds.Count != dependencyIds.Count
                || !dependencyIds.SetEquals(plannedDependencyIds))
            {
                throw new InvalidDataException("The plan must contain exactly one dependency action for every captured dependency.");
            }
            var webPartIds = new HashSet<Guid>(snapshot.WebParts.Select(item => item.Id));
            var plannedWebPartIds = new HashSet<Guid>(plan.WebPartActions.Select(item => item == null ? Guid.Empty : item.SourceWebPartId));
            if (plan.WebPartActions.Any(item => item == null)
                || plan.WebPartActions.Count != webPartIds.Count
                || plannedWebPartIds.Count != webPartIds.Count
                || !webPartIds.SetEquals(plannedWebPartIds))
            {
                throw new InvalidDataException("The plan must contain exactly one Web Part action for every captured shared Web Part.");
            }
        }

        private static string IssueIdentity(MigrationIssue issue)
        {
            if (issue == null)
            {
                return "<null>";
            }
            return issue.Code + "\u001f" + issue.Severity + "\u001f" + issue.Subject + "\u001f"
                + issue.Ingredient + "\u001f" + issue.Message + "\u001f" + issue.SourceIdentity + "\u001f" + issue.TargetIdentity;
        }
    }
}
