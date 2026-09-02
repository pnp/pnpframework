using PnP.Framework.Migration.Pages.Capture;
using PnP.Framework.Migration.Pages.Fields;
using PnP.Framework.Migration.Pages.Lifecycle;
using PnP.Framework.Migration.Pages.Publishing.Packaging;
using PnP.Framework.Migration.Pages.Publishing.Lifecycle;
using PnP.Framework.Migration.Pages.Planning;
using PnP.Framework.Migration.Pages.ClassicWebParts.Bindings;
using System;
using System.IO;
using System.Linq;
using PnP.Framework.Migration.Pages.Cohorts;
using PnP.Framework.Migration.Pages.Runtime;
using PnP.Framework.Migration.Pages.Publishing.Profiles;
using PnP.Framework.Migration.Taxonomy;

namespace PnP.Framework.Migration.Pages.Publishing.Execution
{
    internal static class PublishingPageImportPlanValidator
    {
        public static void Validate(
            PublishingPageMigrationPackage package,
            PublishingPageWorkflowPolicy workflowPolicy)
        {
            if (workflowPolicy == null
                || !string.Equals(package.Selection?.WorkflowId, workflowPolicy.WorkflowId, StringComparison.Ordinal))
            {
                throw new InvalidDataException($"Package workflow '{package.Selection?.WorkflowId}' does not match the selected Publishing Page importer.");
            }

            var expectedSelection = workflowPolicy.Select(package.Snapshot.Source.ContentTypeId);
            if (!string.Equals(
                    PublishingPageDigest.ComputeSelectionDigest(expectedSelection),
                    package.SelectionDigest,
                    StringComparison.OrdinalIgnoreCase))
            {
                throw new InvalidDataException("The sealed validation-cohort assessment does not match the selected workflow policy and source evidence.");
            }

            if (package.Selection?.ValidationCohort?.Disposition != ValidationCohortDisposition.Included)
            {
                throw new InvalidDataException($"The package is not included in validation cohort '{package.Selection?.ValidationCohort?.CohortId}'.");
            }

            if (!string.Equals(package.Snapshot.Runtime?.AdapterId, PageRuntimeAdapterIds.Publishing, StringComparison.Ordinal))
            {
                throw new InvalidDataException($"Runtime adapter '{package.Snapshot.Runtime?.AdapterId ?? PageRuntimeAdapterIds.Unknown}' cannot be imported by the Publishing Page runtime.");
            }

            if (package.Plan.Operation != PageMigrationOperation.CreatePage || !package.Plan.CreateOnly)
            {
                throw new NotSupportedException($"Migration operation '{package.Plan.Operation}' is not executable by this importer.");
            }

            var derivedLifecycle = PublishingPageLifecyclePolicy.DeriveTargetLifecycle(package.Snapshot.Lifecycle);
            if (package.Plan.TargetLifecycle != derivedLifecycle)
            {
                throw new InvalidDataException($"Planned lifecycle '{package.Plan.TargetLifecycle}' does not match the source-derived lifecycle '{derivedLifecycle}'.");
            }

            var fieldByName = package.Snapshot.Fields.ToDictionary(item => item.InternalName, StringComparer.OrdinalIgnoreCase);
            foreach (var action in package.Plan.FieldActions.Where(item => item.Disposition == PageFieldDisposition.Apply))
            {
                if (!string.Equals(action.SourceInternalName, action.TargetInternalName, StringComparison.OrdinalIgnoreCase)
                    || !fieldByName.TryGetValue(action.SourceInternalName, out var field)
                    || field.ReadOnly
                    || !field.HasValue
                    || field.CaptureStatus != PageCaptureStatus.Captured
                    || !PageFieldPlanner.IsImportableKind(field.Kind))
                {
                    throw new InvalidDataException($"Field action '{action.SourceInternalName}' is marked Apply but is not supported by the Publishing Page importer.");
                }
            }
            foreach (var action in package.Plan.FieldActions.Where(item => item.Disposition == PageFieldDisposition.ApplyTaxonomyRelationships))
            {
                if (!string.Equals(action.SourceInternalName, action.TargetInternalName, StringComparison.OrdinalIgnoreCase)
                    || !fieldByName.TryGetValue(action.SourceInternalName, out var field)
                    || field.ReadOnly
                    || !field.HasValue
                    || field.CaptureStatus != PageCaptureStatus.Captured
                    || !PageFieldPlanner.IsTaxonomy(field.Kind))
                {
                    throw new InvalidDataException($"Field action '{action.SourceInternalName}' is marked for taxonomy replay but is not supported by the Publishing Page importer.");
                }
                var relationshipActions = package.Plan.TaxonomyRelationshipActions
                    .Where(value => value.SourceFieldId == field.Id)
                    .ToArray();
                if (relationshipActions.Length != field.TaxonomyValues.Count
                    || relationshipActions.Any(value => !value.IsExecutable))
                {
                    throw new InvalidDataException($"Field action '{action.SourceInternalName}' has incomplete or blocked taxonomy relationship actions.");
                }
            }

            if (package.Plan.IsExecutable
                && ((package.Plan.Topology != null
                        && (package.Plan.TopologyTargetAnalysis == null || !package.Plan.TopologyTargetAnalysis.IsAdmitted))
                    || (package.Snapshot.ListDependencies.Count > 0
                        && (package.Plan.ListMigration == null || !package.Plan.ListMigration.IsExecutable))
                    || package.Plan.WebPartActions.Any(value => value.Disposition == ClassicWebPartDisposition.Block)))
            {
                throw new InvalidDataException("An executable publishing-page plan contains a blocked List or Web Part action.");
            }
        }
    }
}
