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
using PnP.Framework.Migration.Pages.Publishing.Planning;
using PnP.Framework.Migration.Pages.Publishing.Reporting;
using PnP.Framework.Migration.Pages.Publishing.Verification;
using PnP.Framework.Migration.Verification;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;

namespace PnP.Framework.Migration.Pages.Publishing.EnterpriseWiki
{
    public sealed class EnterpriseWikiMigrationPlanner
    {
        private static readonly RuntimeVerificationRequirement[] RuntimeVerificationRequirements =
        {
            Requirement("page-reachability", RuntimeVerificationRequirementKind.PageReachability, "Fresh navigation reaches the target classic page."),
            Requirement("error-shell-absence", RuntimeVerificationRequirementKind.ErrorShellAbsence, "The target is not a login, access denied, not found, or SharePoint error shell."),
            Requirement("authored-dom-equality", RuntimeVerificationRequirementKind.AuthoredDomEquality, "Normalized authored DOM is equal."),
            Requirement("resource-inventory-equality", RuntimeVerificationRequirementKind.ResourceInventoryEquality, "Authored resource inventory is equal."),
            Requirement("script-inventory-equality", RuntimeVerificationRequirementKind.ScriptInventoryEquality, "Authored script inventory is equal."),
            Requirement("inline-event-inventory-equality", RuntimeVerificationRequirementKind.InlineEventInventoryEquality, "Inline event inventory is equal."),
            Requirement("screenshot-capture", RuntimeVerificationRequirementKind.ScreenshotCapture, "Full-page and authored-canvas screenshots are captured.")
        };

        public PublishingPageMigrationPackage Plan(
            ClientContext targetContext,
            PublishingPageExportPackage exportPackage,
            PagePlanningOptions options)
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
            var targetPagePath = PagePath.Normalize(targetWeb.ServerRelativeUrl, options.TargetPageServerRelativeUrl, "Pages");
            var blockers = snapshot.Blockers.ToList();
            var warnings = snapshot.Warnings.ToList();
            AddSecurityPolicyDecision(snapshot, options, blockers);
            AddManagedMetadataDecision(snapshot, options, blockers);

            var targetLifecycle = PublishingPageLifecyclePolicy.DeriveTargetLifecycle(snapshot.Lifecycle);
            var lifecycleReason = DescribeLifecycleDecision(snapshot.Lifecycle, targetLifecycle, warnings);
            var replacements = PageReferencePlanner.BuildTextReplacements(snapshot.Source, targetWeb.Url, targetWeb.ServerRelativeUrl);
            var dependencyActions = PageReferencePlanner.BuildActions(
                snapshot.Source,
                snapshot.Dependencies,
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
            var expectedContent = PageTextTransformer.Rewrite(snapshot.PublishingPageContent, replacements);
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
                Operation = PageMigrationOperation.CreatePage,
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
                RuntimeVerification = new RuntimeVerificationManifest
                {
                    Requirements = RuntimeVerificationRequirements.Select(CopyRequirement).ToList()
                },
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

        private static RuntimeVerificationRequirement Requirement(
            string id,
            RuntimeVerificationRequirementKind kind,
            string description)
        {
            return new RuntimeVerificationRequirement
            {
                Id = id,
                Kind = kind,
                Description = description,
                Required = true
            };
        }

        private static RuntimeVerificationRequirement CopyRequirement(RuntimeVerificationRequirement requirement)
        {
            return new RuntimeVerificationRequirement
            {
                Id = requirement.Id,
                Kind = requirement.Kind,
                Required = requirement.Required,
                Description = requirement.Description
            };
        }

        private static void ValidateOptions(PagePlanningOptions options)
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
            PagePlanningOptions options,
            ICollection<string> blockers)
        {
            if (snapshot.Security.HasUniqueRoleAssignments && options.RequireInheritedPermissions)
            {
                blockers.Add("The source page has unique role assignments. The Enterprise Wiki profile requires inherited permissions.");
            }
        }

        private static void AddManagedMetadataDecision(
            PublishingPageCaptureBundle snapshot,
            PagePlanningOptions options,
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

        private static PagePlanningOptions CopyOptions(
            PagePlanningOptions options,
            string targetPagePath)
        {
            return new PagePlanningOptions
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
