using PnP.Framework.Migration.Pages.Capture;
using PnP.Framework.Migration.Pages.Content;
using PnP.Framework.Migration.Pages.Lifecycle;
using PnP.Framework.Migration.Pages.Planning;
using PnP.Framework.Migration.Pages.Publishing.Capture;
using PnP.Framework.Migration.Pages.Publishing.Layouts.Packaging;
using PnP.Framework.Migration.Pages.Publishing.Lifecycle;
using PnP.Framework.Migration.Pages.Publishing.Planning;
using PnP.Framework.Migration.Packaging;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace PnP.Framework.Migration.Pages.Publishing.Packaging
{
    public static class PublishingPagePackageValidator
    {
        public static void ValidateExport(PublishingPageExportPackage package)
        {
            ValidateExport(package, null);
        }

        public static void ValidateExport(
            PublishingPageExportPackage package,
            IMigrationArtifactStore artifactStore)
        {
            if (package == null)
            {
                throw new InvalidDataException("The publishing-page export is empty.");
            }

            if (!string.Equals(package.SchemaVersion, PublishingPagePackageContract.ExportSchemaVersion, StringComparison.Ordinal))
            {
                throw new InvalidDataException($"Unsupported publishing-page export schema '{package.SchemaVersion}'.");
            }

            var snapshot = package.Snapshot;
            if (snapshot == null)
            {
                throw new InvalidDataException("The publishing-page export must contain a source snapshot.");
            }

            if (string.IsNullOrWhiteSpace(snapshot.SourceProfile)
                || snapshot.Source == null
                || snapshot.Layout == null
                || snapshot.CapturePolicy == null
                || snapshot.Security == null
                || snapshot.Lifecycle == null
                || snapshot.SourceFence == null)
            {
                throw new InvalidDataException("The source snapshot is missing its profile, identity, publishing layout, policy, security, lifecycle, or source fence.");
            }

            if (snapshot.Fields == null
                || snapshot.WebParts == null
                || snapshot.Dependencies == null
                || snapshot.Blockers == null
                || snapshot.Warnings == null)
            {
                throw new InvalidDataException("The source snapshot contains a null inventory collection.");
            }

            var contentDigest = PublishingPageDigest.ComputeSha256(snapshot.PublishingPageContent ?? string.Empty);
            if (!string.Equals(contentDigest, snapshot.PublishingPageContentSha256, StringComparison.OrdinalIgnoreCase))
            {
                throw new InvalidDataException("The PublishingPageContent digest does not match the source HTML.");
            }

            PublishingPageLayoutPackageValidator.ValidateSnapshot(snapshot.Layout, artifactStore);

            var duplicateField = snapshot.Fields
                .GroupBy(item => item?.InternalName, StringComparer.OrdinalIgnoreCase)
                .FirstOrDefault(group => string.IsNullOrWhiteSpace(group.Key) || group.Count() > 1);
            if (duplicateField != null)
            {
                throw new InvalidDataException($"The source field inventory contains a missing or duplicate internal name '{duplicateField.Key}'.");
            }

            foreach (var webPart in snapshot.WebParts)
            {
                if (webPart == null)
                {
                    throw new InvalidDataException("The Web Part inventory contains a null entry.");
                }

                if (!string.Equals(
                        PublishingPageDigest.ComputeSha256(webPart.ExportXml ?? string.Empty),
                        webPart.ExportSha256,
                        StringComparison.OrdinalIgnoreCase))
                {
                    throw new InvalidDataException($"Web Part export digest mismatch: {webPart.Id}");
                }
            }

            var duplicateDependency = snapshot.Dependencies
                .GroupBy(item => item?.Id, StringComparer.Ordinal)
                .FirstOrDefault(group => string.IsNullOrWhiteSpace(group.Key) || group.Count() > 1);
            if (duplicateDependency != null)
            {
                throw new InvalidDataException($"The dependency inventory contains a missing or duplicate ID '{duplicateDependency.Key}'.");
            }

            foreach (var dependency in snapshot.Dependencies.Where(item => !string.IsNullOrWhiteSpace(item.ContentBase64)))
            {
                byte[] payload;
                try
                {
                    payload = Convert.FromBase64String(dependency.ContentBase64);
                }
                catch (FormatException exception)
                {
                    throw new InvalidDataException($"Dependency payload is not valid Base64: {dependency.Id}", exception);
                }

                if (payload.LongLength != dependency.ContentLength
                    || !string.Equals(
                        PublishingPageDigest.ComputeSha256(payload),
                        dependency.ContentSha256,
                        StringComparison.OrdinalIgnoreCase))
                {
                    throw new InvalidDataException($"Dependency payload length or digest mismatch: {dependency.Id}");
                }
            }

            var snapshotDigest = PublishingPageDigest.ComputeSnapshotDigest(snapshot);
            if (!string.Equals(snapshotDigest, package.SnapshotDigest, StringComparison.OrdinalIgnoreCase))
            {
                throw new InvalidDataException("The source snapshot digest does not match the export payload.");
            }
        }

        public static void ValidateMigration(PublishingPageMigrationPackage package)
        {
            ValidateMigration(package, null);
        }

        public static void ValidateMigration(
            PublishingPageMigrationPackage package,
            IMigrationArtifactStore artifactStore)
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

            ValidateExport(new PublishingPageExportPackage
            {
                SchemaVersion = package.ExportSchemaVersion,
                ExportedAtUtc = package.ExportedAtUtc,
                Snapshot = package.Snapshot,
                SnapshotDigest = package.SnapshotDigest
            }, artifactStore);

            var plan = package.Plan;
            if (plan.PlanningPolicy == null
                || plan.TargetProbe == null
                || plan.LayoutMaterialization == null
                || plan.LayoutAdmission == null
                || plan.FieldActions == null
                || plan.DependencyActions == null
                || plan.Replacements == null
                || plan.StorageAssertions == null
                || plan.RuntimeVerification == null
                || plan.RuntimeVerification.Requirements == null
                || plan.Blockers == null
                || plan.Warnings == null)
            {
                throw new InvalidDataException("The migration plan is missing policy, target probe, or an action/assertion collection.");
            }

            if (plan.PlanningPolicy.TaxonomySchemaMappings == null)
            {
                throw new InvalidDataException("The planning policy contains a null taxonomy schema mapping collection.");
            }

            PublishingPageLayoutPackageValidator.ValidatePlan(
                plan.PageLayoutName,
                plan.IsExecutable,
                plan.LayoutMaterialization,
                plan.LayoutTargetProbe,
                plan.LayoutAdmission);

            var duplicateRuntimeRequirement = plan.RuntimeVerification.Requirements
                .GroupBy(item => item?.Id, StringComparer.Ordinal)
                .FirstOrDefault(group => string.IsNullOrWhiteSpace(group.Key) || group.Count() > 1);
            if (duplicateRuntimeRequirement != null || plan.RuntimeVerification.Requirements.Any(item => item == null))
            {
                throw new InvalidDataException($"The runtime verification manifest contains a missing or duplicate requirement ID '{duplicateRuntimeRequirement?.Key}'.");
            }

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
        }

    }
}
