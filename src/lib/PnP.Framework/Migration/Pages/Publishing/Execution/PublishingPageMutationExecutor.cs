using Microsoft.SharePoint.Client;
using PnP.Framework.Migration.Execution;
using PnP.Framework.Migration.Pages.References;
using PnP.Framework.Migration.Pages.Publishing.Packaging;
using PnP.Framework.Migration.Pages.Publishing.Layouts;
using PnP.Framework.Migration.Pages.Publishing.Verification;
using PnP.Framework.Migration.Packaging;
using PnP.Framework.Migration.Schema.ContentTypes;
using PnP.Framework.Migration.Lists.Execution;
using PnP.Framework.Migration.Lists.Planning;
using System;
using System.Collections.Generic;
using System.Linq;

namespace PnP.Framework.Migration.Pages.Publishing.Execution
{
    internal static class PublishingPageMutationExecutor
    {
        public static PublishingPageImportReceipt Execute(
            ClientContext targetContext,
            PublishingPageMigrationPackage package,
            string approvedPlanDigest,
            Guid operationId,
            DateTimeOffset startedAt,
            MigrationExecutionRecorder recorder,
            IMigrationArtifactStore artifactStore,
            Func<string, bool> isExpectedContentType)
        {
            var warnings = new List<string>();
            MaterializeLayout(targetContext, package, recorder, artifactStore);
            var materializedDependencies = MaterializeDependencies(targetContext, package, recorder);
            var listReceipts = ListMaterializationCoordinator.Ensure(
                targetContext,
                package.Snapshot.ListDependencies,
                package.Plan.ListMigration,
                recorder,
                artifactStore);
            var writeResult = PublishingPageTargetWriter.Write(targetContext, package, listReceipts, recorder, warnings);
            PublishingPageLifecycleApplier.Apply(targetContext, package, writeResult, recorder, warnings);
            var receipt = PublishingPageImportVerifier.Verify(
                targetContext,
                package,
                approvedPlanDigest,
                operationId,
                startedAt,
                materializedDependencies,
                listReceipts.Values.OrderBy(value => value.SourceListId).ToList(),
                writeResult.FieldResults,
                recorder.Steps,
                warnings,
                isExpectedContentType);
            recorder.RecordState(receipt.ExecutionStatus, receipt.FreshReadbackPassed
                ? "Mutation and fresh storage verification completed."
                : "Mutation completed, but fresh storage verification failed.");
            return receipt;
        }

        private static void MaterializeLayout(
            ClientContext context,
            PublishingPageMigrationPackage package,
            MigrationExecutionRecorder recorder,
            IMigrationArtifactStore artifactStore)
        {
            var plan = package.Plan.LayoutMaterialization;
            var admission = package.Plan.LayoutAdmission;
            if (admission.Disposition == PublishingPageLayoutMaterializationDisposition.ReuseTargetStock)
            {
                recorder.RecordAlreadySatisfied("layout.schema", "The approved plan reuses target-runtime Page Layout schema.");
                recorder.RecordAlreadySatisfied("layout.resources", "The approved plan reuses target-runtime Page Layout resources.");
                recorder.RecordAlreadySatisfied("layout.file", $"Reuse stock Page Layout '{plan.TargetServerRelativeUrl}'.");
                PublishingPageLayoutMaterializer.Ensure(
                    context,
                    package.Snapshot.Layout,
                    plan,
                    admission,
                    artifactStore);
                return;
            }

            recorder.Execute(
                "layout.schema",
                $"Ensure associated content type closure '{plan.AssociatedContentTypeName}' ({plan.AssociatedContentTypeId}).",
                () =>
                {
                    return PublishingPageLayoutSchemaMaterializer.Ensure(context, plan, admission);
                },
                disposition => disposition == ContentTypeMaterializationDisposition.ReuseOwned
                    ? MutationOutcome.AlreadySatisfied
                    : MutationOutcome.Applied,
                disposition => $"Associated content type schema disposition: {disposition}.");

            var plannedResources = plan.ResourceMaterializations.Count(value =>
                value.Disposition == PublishingPageLayoutResourceMaterializationDisposition.CreateOrReuseOwned);
            if (plannedResources == 0)
            {
                recorder.RecordAlreadySatisfied("layout.resources", "The approved layout has no owned rendering resources to materialize.");
            }
            else
            {
                recorder.Execute(
                    "layout.resources",
                    $"Ensure {plannedResources} approved Page Layout rendering resource(s).",
                    () => PublishingPageLayoutResourceMaterializer.Ensure(context, plan, artifactStore),
                    count => count == 0 ? MutationOutcome.AlreadySatisfied : MutationOutcome.Applied,
                    count => $"Created {count} Page Layout rendering resource(s); exact existing resources were reused.");
            }

            recorder.Execute(
                "layout.file",
                $"Ensure Page Layout '{plan.TargetServerRelativeUrl}'.",
                () => PublishingPageLayoutMaterializer.Ensure(
                    context,
                    package.Snapshot.Layout,
                    plan,
                    admission,
                    artifactStore),
                created => created ? MutationOutcome.Applied : MutationOutcome.AlreadySatisfied,
                created => created ? "Created and freshly verified the digest-owned Page Layout." : "Reused and freshly verified the exact Page Layout.");
        }

        private static int MaterializeDependencies(
            ClientContext context,
            PublishingPageMigrationPackage package,
            MigrationExecutionRecorder recorder)
        {
            var actionCount = package.Plan.DependencyActions.Count(action => action.Disposition == PageReferenceDisposition.MaterializeAtTarget);
            if (actionCount == 0)
            {
                recorder.RecordAlreadySatisfied("dependencies.materialize", "The approved plan has no dependency artifacts to materialize.");
                return 0;
            }

            return recorder.Execute(
                "dependencies.materialize",
                $"Materialize {actionCount} approved dependency artifact(s).",
                () => PageReferenceMaterializer.Materialize(context, package.Snapshot.Dependencies, package.Plan.DependencyActions),
                count => count == 0 ? MutationOutcome.AlreadySatisfied : MutationOutcome.Applied,
                count => $"Materialized {count} dependency artifact(s).");
        }
    }
}
