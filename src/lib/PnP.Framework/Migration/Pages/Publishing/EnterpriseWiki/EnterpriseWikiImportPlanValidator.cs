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

namespace PnP.Framework.Migration.Pages.Publishing.EnterpriseWiki
{
    internal static class EnterpriseWikiImportPlanValidator
    {
        public static void Validate(PublishingPageMigrationPackage package)
        {
            if (!string.Equals(package.Snapshot.SourceProfile, EnterpriseWikiMigrationProfile.SourceProfile, StringComparison.Ordinal))
            {
                throw new InvalidDataException($"The '{package.Snapshot.SourceProfile}' source profile cannot be imported by the Enterprise Wiki importer.");
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
                if (!EnterpriseWikiMigrationProfile.AdditionalFieldNames.Contains(action.SourceInternalName)
                    || !string.Equals(action.SourceInternalName, action.TargetInternalName, StringComparison.OrdinalIgnoreCase)
                    || !fieldByName.TryGetValue(action.SourceInternalName, out var field)
                    || field.ReadOnly
                    || !field.HasValue
                    || field.CaptureStatus != PageCaptureStatus.Captured
                    || !PageFieldPlanner.IsImportableKind(field.Kind))
                {
                    throw new InvalidDataException($"Field action '{action.SourceInternalName}' is marked Apply but is not supported by the Enterprise Wiki importer.");
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
