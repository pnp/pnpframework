using Microsoft.SharePoint.Client;
using PnP.Framework.Migration.Execution;
using PnP.Framework.Migration.Pages.References;
using PnP.Framework.Migration.Pages.Publishing.Packaging;
using PnP.Framework.Migration.Pages.Publishing.Verification;
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
            Func<string, bool> isExpectedContentType)
        {
            var warnings = new List<string>();
            var materializedDependencies = MaterializeDependencies(targetContext, package, recorder);
            var writeResult = PublishingPageTargetWriter.Write(targetContext, package, recorder, warnings);
            PublishingPageLifecycleApplier.Apply(targetContext, package, writeResult, recorder, warnings);
            var receipt = PublishingPageImportVerifier.Verify(
                targetContext,
                package,
                approvedPlanDigest,
                operationId,
                startedAt,
                materializedDependencies,
                writeResult.FieldResults,
                recorder.Steps,
                warnings,
                isExpectedContentType);
            recorder.RecordState(receipt.ExecutionStatus, receipt.FreshReadbackPassed
                ? "Mutation and fresh storage verification completed."
                : "Mutation completed, but fresh storage verification failed.");
            return receipt;
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
