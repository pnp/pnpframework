using Microsoft.SharePoint.Client;
using PnP.Framework.Migration.PublishingPages.Capture;
using PnP.Framework.Migration.PublishingPages.Content;
using PnP.Framework.Migration.PublishingPages.Fields;
using PnP.Framework.Migration.PublishingPages.Lifecycle;
using PnP.Framework.Migration.PublishingPages.Packaging;
using PnP.Framework.Migration.PublishingPages.Planning;
using PnP.Framework.Migration.PublishingPages.References;
using PnP.Framework.Migration.PublishingPages.Reporting;
using PnP.Framework.Migration.PublishingPages.Verification;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;

namespace PnP.Framework.Migration.PublishingPages.EnterpriseWiki
{
    public sealed class EnterpriseWikiMigrationPlanner
    {
        private static readonly string[] BrowserAssertions =
        {
            "fresh-navigation-reaches-target-classic-page",
            "no-login-access-denied-not-found-or-sharepoint-error-shell",
            "normalized-authored-dom-equal",
            "resource-script-and-inline-event-inventory-equal",
            "full-page-and-authored-canvas-screenshots-captured"
        };

        public PublishingPageMigrationPackage Plan(
            ClientContext targetContext,
            PublishingPageExportPackage exportPackage,
            PublishingPagePlanningOptions options)
        {
            if (targetContext == null)
            {
                throw new ArgumentNullException(nameof(targetContext));
            }

            PublishingPagePackageValidator.ValidateExport(exportPackage);
            ValidateOptions(options);
            if (!string.Equals(exportPackage.Snapshot.SourceProfile, EnterpriseWikiMigrationProfile.SourceProfile, StringComparison.Ordinal))
            {
                throw new InvalidOperationException($"The '{exportPackage.Snapshot.SourceProfile}' source profile cannot be planned by the Enterprise Wiki planner.");
            }

            var targetWeb = targetContext.Web;
            targetContext.Load(targetWeb, web => web.Url, web => web.ServerRelativeUrl);
            targetContext.ExecuteQueryRetry();

            var snapshot = exportPackage.Snapshot;
            var targetPagePath = PublishingPagePath.Normalize(targetWeb.ServerRelativeUrl, options.TargetPageServerRelativeUrl, "Pages");
            var blockers = snapshot.Blockers.ToList();
            var warnings = snapshot.Warnings.ToList();
            AddSecurityPolicyDecision(snapshot, options, blockers);
            AddManagedMetadataDecision(snapshot, options, blockers);

            var targetLifecycle = PageLifecyclePolicy.DeriveTargetLifecycle(snapshot.Lifecycle);
            var lifecycleReason = DescribeLifecycleDecision(snapshot.Lifecycle, targetLifecycle, warnings);
            var replacements = PageReferencePlanner.BuildTextReplacements(snapshot.Source, targetWeb.Url, targetWeb.ServerRelativeUrl);
            var dependencyActions = PageReferencePlanner.BuildActions(
                snapshot,
                targetWeb.Url,
                targetWeb.ServerRelativeUrl,
                options,
                blockers);
            var targetProbe = EnterpriseWikiTargetInspector.Inspect(
                targetContext,
                targetPagePath,
                dependencyActions,
                targetLifecycle,
                blockers);
            var fieldActions = PageFieldPlanner.BuildActions(
                targetContext,
                snapshot.Fields,
                EnterpriseWikiMigrationProfile.HandledFieldNames,
                EnterpriseWikiMigrationProfile.AdditionalFieldNames,
                options,
                warnings);
            var expectedContent = PublishingPageContentTransformer.Rewrite(snapshot.PublishingPageContent, replacements);
            var expectedContentDigest = PublishingPageDigest.ComputeSha256(expectedContent);
            var plan = new PublishingPageMigrationPlan
            {
                SourceSnapshotDigest = exportPackage.SnapshotDigest,
                SourceWebUrl = snapshot.Source.WebUrl,
                SourcePageServerRelativeUrl = snapshot.Source.PageServerRelativeUrl,
                TargetWebUrl = targetWeb.Url.TrimEnd('/'),
                TargetWebServerRelativeUrl = targetWeb.ServerRelativeUrl,
                TargetPageServerRelativeUrl = targetPagePath,
                PageLayoutName = EnterpriseWikiMigrationProfile.PageLayoutName,
                Operation = PublishingPageMigrationOperation.CreatePage,
                TargetLifecycle = targetLifecycle,
                LifecycleReason = lifecycleReason,
                CreateOnly = options.CreateOnly,
                PlanningPolicy = CopyOptions(options, targetPagePath),
                TargetProbe = targetProbe,
                FieldActions = fieldActions,
                DependencyActions = dependencyActions,
                Replacements = replacements,
                ExpectedPublishingPageContentSha256 = expectedContentDigest,
                StorageAssertions = PageStorageAssertionBuilder.Build(
                    snapshot,
                    targetPagePath,
                    dependencyActions,
                    expectedContentDigest,
                    targetLifecycle),
                BrowserAssertions = BrowserAssertions.ToList(),
                Blockers = blockers.Distinct(StringComparer.Ordinal).OrderBy(item => item, StringComparer.Ordinal).ToList(),
                Warnings = warnings.Distinct(StringComparer.Ordinal).OrderBy(item => item, StringComparer.Ordinal).ToList()
            };
            var package = new PublishingPageMigrationPackage
            {
                PlannedAtUtc = DateTimeOffset.UtcNow,
                ExportedAtUtc = exportPackage.ExportedAtUtc,
                State = plan.IsExecutable ? PublishingPagePackageState.ApprovalReady : PublishingPagePackageState.Blocked,
                Snapshot = snapshot,
                Plan = plan,
                SnapshotDigest = exportPackage.SnapshotDigest,
                PlanDigest = PublishingPageDigest.ComputePlanDigest(plan),
                Report = BuildReportSummary(snapshot, plan)
            };
            PublishingPagePackageValidator.ValidateMigration(package);
            return package;
        }

        private static void ValidateOptions(PublishingPagePlanningOptions options)
        {
            if (options == null)
            {
                throw new ArgumentNullException(nameof(options));
            }

            if (string.IsNullOrWhiteSpace(options.TargetPageServerRelativeUrl))
            {
                throw new ArgumentException("A target page path is required.", nameof(options));
            }

            if (!options.CreateOnly)
            {
                throw new NotSupportedException("Only create-page plans are supported. Deferred-field recovery remains represented by the package schema but is not executable yet.");
            }
        }

        private static void AddSecurityPolicyDecision(
            PublishingPageCaptureBundle snapshot,
            PublishingPagePlanningOptions options,
            ICollection<string> blockers)
        {
            if (snapshot.Security.HasUniqueRoleAssignments && options.RequireInheritedPermissions)
            {
                blockers.Add("The source page has unique role assignments. The Enterprise Wiki profile requires inherited permissions.");
            }
        }

        private static void AddManagedMetadataDecision(
            PublishingPageCaptureBundle snapshot,
            PublishingPagePlanningOptions options,
            ICollection<string> blockers)
        {
            var managedMetadata = snapshot.Fields
                .Where(field => field.HasValue)
                .Where(field => EnterpriseWikiMigrationProfile.AdditionalFieldNames.Contains(field.InternalName))
                .Where(field => field.Kind == PageFieldValueKind.Taxonomy || field.Kind == PageFieldValueKind.TaxonomyCollection)
                .ToArray();
            if (managedMetadata.Length > 0 && options.BlockOnManagedMetadata)
            {
                blockers.Add($"The source snapshot contains {managedMetadata.Length} non-empty managed metadata field value(s), but no reviewed target term mapping was supplied.");
            }
        }

        private static string DescribeLifecycleDecision(
            PageLifecycleSnapshot lifecycle,
            PublishingPageTargetLifecycle targetLifecycle,
            ICollection<string> warnings)
        {
            if (lifecycle == null || string.IsNullOrWhiteSpace(lifecycle.Level))
            {
                warnings.Add("Source lifecycle evidence is incomplete. The conservative target lifecycle is Draft.");
                return "Source lifecycle evidence is incomplete, so the target will remain Draft.";
            }

            if (string.Equals(lifecycle.Level, "Published", StringComparison.OrdinalIgnoreCase)
                && targetLifecycle == PublishingPageTargetLifecycle.Draft)
            {
                warnings.Add("Source lifecycle evidence is contradictory. The conservative target lifecycle is Draft.");
                return $"The source reports Published but has conflicting checkout '{lifecycle.CheckOutType ?? "unknown"}' or moderation '{lifecycle.ModerationStatus?.ToString(CultureInfo.InvariantCulture) ?? "unknown"}' evidence, so the target will remain Draft.";
            }

            return targetLifecycle == PublishingPageTargetLifecycle.Published
                ? "The source file level is Published with no conflicting checkout or moderation evidence, so the target will be published."
                : $"The source file level is '{lifecycle.Level}', so the target will remain Draft.";
        }

        private static PublishingPagePlanningOptions CopyOptions(
            PublishingPagePlanningOptions options,
            string targetPagePath)
        {
            return new PublishingPagePlanningOptions
            {
                TargetPageServerRelativeUrl = targetPagePath,
                RequireInheritedPermissions = options.RequireInheritedPermissions,
                BlockOnManagedMetadata = options.BlockOnManagedMetadata,
                AllowExternalResourceReferences = options.AllowExternalResourceReferences,
                CreateOnly = options.CreateOnly
            };
        }

        private static PublishingPageMigrationReport BuildReportSummary(
            PublishingPageCaptureBundle snapshot,
            PublishingPageMigrationPlan plan)
        {
            return new PublishingPageMigrationReport
            {
                Summary = plan.IsExecutable
                    ? "Source export and target analysis completed. Import requires explicit approval of the sealed plan digest."
                    : "The package is sealed for review but cannot be imported until every blocker is resolved and a new plan is generated.",
                CapturedIngredients = new List<string>
                {
                    "Page/file/list item identity and source stability fence",
                    $"{snapshot.SourceProfile} content type and publishing layout evidence",
                    $"All {snapshot.Fields.Count} source Pages-library field definitions and returned values",
                    $"{snapshot.WebParts.Count} shared Web Part export(s) with zone placement",
                    $"{snapshot.Dependencies.Count} authored dependency/link snapshot(s)",
                    "Page security inheritance and source lifecycle evidence",
                    "Target publishing library, versioning, lifecycle, field, layout, and create-only probes"
                },
                Blockers = plan.Blockers.ToList(),
                Warnings = plan.Warnings.ToList()
            };
        }
    }
}
