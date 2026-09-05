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
using PnP.Framework.Migration.Topology.Execution;
using PnP.Framework.Migration.Topology;
using PnP.Framework.Migration.Schema.Fields;
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
            PublishingPageExecutionScope executionScope,
            string approvedPlanDigest,
            Guid operationId,
            DateTimeOffset startedAt,
            MigrationExecutionRecorder recorder,
            IMigrationArtifactStore artifactStore,
            string expectedContentTypeId)
        {
            var warnings = new List<string>();
            var topologyReceipt = TopologyMaterializationCoordinator.Ensure(
                targetContext,
                package.Plan.Topology,
                executionScope.TopologyPlan,
                recorder);
            MaterializeLayout(targetContext, package, executionScope, recorder, artifactStore);
            var materializedDependencies = MaterializeDependencies(
                targetContext,
                package,
                executionScope,
                topologyReceipt,
                recorder);
            var listReceipts = ListMaterializationCoordinator.Ensure(
                targetContext,
                package.Snapshot.ListDependencies,
                package.Plan.ListMigration,
                executionScope.ListScope,
                recorder,
                artifactStore);
            PublishingPageWriteResult writeResult = null;
            if (executionScope.PageArtifact)
            {
                writeResult = PublishingPageTargetWriter.Write(
                    targetContext,
                    package,
                    executionScope,
                    listReceipts,
                    recorder,
                    warnings);
                if (writeResult.ResumedExistingOwnedPage)
                {
                    recorder.RecordAlreadySatisfied(
                        "page.checkin",
                        "The exact migration-owned page is already checked in; fresh readback will verify storage and lifecycle without creating another version.");
                    recorder.RecordAlreadySatisfied(
                        "page.publish",
                        "The exact migration-owned page already exists; fresh readback will verify the approved lifecycle without republishing it.");
                    recorder.RecordAlreadySatisfied(
                        "page.approve",
                        "The exact migration-owned page already exists; fresh readback will verify moderation state without another approval mutation.");
                }
                else
                {
                    PublishingPageLifecycleApplier.Apply(
                        targetContext,
                        package,
                        writeResult,
                        executionScope.Lifecycle,
                        recorder,
                        warnings);
                }
            }
            else
            {
                recorder.RecordAlreadySatisfied(
                    "page.create",
                    "The Page artifact is outside the admitted execution frontier; independent ingredients were materialized without creating a page.");
            }
            var receipt = PublishingPageImportVerifier.Verify(
                targetContext,
                package,
                executionScope,
                approvedPlanDigest,
                operationId,
                startedAt,
                materializedDependencies,
                topologyReceipt,
                listReceipts.Values.OrderBy(value => value.SourceListId).ToList(),
                writeResult?.FieldResults ?? Array.Empty<PnP.Framework.Migration.Pages.Fields.PageFieldImportResult>(),
                recorder.Steps,
                warnings,
                expectedContentTypeId);
            recorder.RecordState(receipt.ExecutionStatus, receipt.FreshReadbackPassed
                ? "Mutation and fresh storage verification completed."
                : "Mutation completed, but fresh storage verification failed.");
            return receipt;
        }

        private static void MaterializeLayout(
            ClientContext context,
            PublishingPageMigrationPackage package,
            PublishingPageExecutionScope executionScope,
            MigrationExecutionRecorder recorder,
            IMigrationArtifactStore artifactStore)
        {
            var plan = package.Plan.LayoutMaterialization;
            var admission = package.Plan.LayoutAdmission;
            var selectedFields = executionScope.PageContentTypeFields(package);
            if (!executionScope.ContentType && selectedFields.Count > 0)
            {
                recorder.Execute(
                    "layout.schema.fields",
                    $"Ensure {selectedFields.Count} independently executable Page Content Type field(s).",
                    () => SiteFieldMaterializer.Ensure(context, context.Site.RootWeb, selectedFields),
                    count => count == 0 ? MutationOutcome.AlreadySatisfied : MutationOutcome.Applied,
                    count => $"Created {count} Page Content Type field(s); every selected field passed fresh schema readback.");
            }
            else if (!executionScope.ContentType)
            {
                recorder.RecordAlreadySatisfied(
                    "layout.schema.fields",
                    "No Page Content Type field is present in the admitted execution frontier.");
            }

            var selectedResources = executionScope.LayoutResources(package);
            if (selectedResources.Count > 0)
            {
                var resourcePlan = new PublishingPageLayoutMaterializationPlan
                {
                    ResourceMaterializations = selectedResources
                };
                recorder.Execute(
                    "layout.resources",
                    $"Ensure {selectedResources.Count} independently executable Page Layout rendering resource(s).",
                    () => PublishingPageLayoutResourceMaterializer.Ensure(context, resourcePlan, artifactStore),
                    count => count == 0 ? MutationOutcome.AlreadySatisfied : MutationOutcome.Applied,
                    count => $"Created {count} Page Layout rendering resource(s); exact existing resources were reused.");
            }
            else
            {
                recorder.RecordAlreadySatisfied(
                    "layout.resources",
                    "No Page Layout rendering resource is present in the admitted execution frontier.");
            }

            if (!executionScope.ContentType)
            {
                recorder.RecordAlreadySatisfied(
                    "layout.schema",
                    "The Page Content Type transaction is outside the admitted execution frontier.");
            }
            else
            {
                recorder.Execute(
                    "layout.schema",
                    $"Ensure associated content type closure '{plan.AssociatedContentTypeName}' ({plan.AssociatedContentTypeId}).",
                    () => PublishingPageLayoutSchemaMaterializer.Ensure(context, plan, admission),
                    disposition => disposition == ContentTypeMaterializationDisposition.ReuseOwned
                        ? MutationOutcome.AlreadySatisfied
                        : MutationOutcome.Applied,
                    disposition => $"Associated content type schema disposition: {disposition}.");
            }

            if (!executionScope.Layout)
            {
                recorder.RecordAlreadySatisfied(
                    "layout.file",
                    "The Page Layout transaction is outside the admitted execution frontier.");
                return;
            }
            if (admission.Disposition == PublishingPageLayoutMaterializationDisposition.ReuseTargetStock)
            {
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
            PublishingPageExecutionScope executionScope,
            TopologyMaterializationReceipt topologyReceipt,
            MigrationExecutionRecorder recorder)
        {
            var actions = executionScope.ReferenceActions(package);
            var actionCount = actions.Count(action => action.Disposition == PageReferenceDisposition.MaterializeAtTarget);
            if (actionCount == 0)
            {
                recorder.RecordAlreadySatisfied("dependencies.materialize", "The approved plan has no dependency artifacts to materialize.");
                return 0;
            }

            return recorder.Execute(
                "dependencies.materialize",
                $"Materialize {actionCount} approved dependency artifact(s).",
                () => PageReferenceMaterializer.Materialize(
                    context,
                    package.Snapshot.Dependencies,
                    actions,
                    topologyReceipt),
                count => count == 0 ? MutationOutcome.AlreadySatisfied : MutationOutcome.Applied,
                count => $"Materialized {count} dependency artifact(s).");
        }
    }
}
