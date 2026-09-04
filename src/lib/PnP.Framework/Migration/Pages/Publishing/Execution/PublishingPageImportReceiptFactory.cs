using PnP.Framework.Migration.Execution;
using PnP.Framework.Migration.Pages.Publishing.Packaging;
using PnP.Framework.Migration.Pages.Ingredients;
using PnP.Framework.Migration.Verification;
using System;
using System.Collections.Generic;
using System.Linq;

namespace PnP.Framework.Migration.Pages.Publishing.Execution
{
    internal static class PublishingPageImportReceiptFactory
    {
        public static PublishingPageImportReceipt AdmissionFailure(
            PublishingPageMigrationPackage package,
            Guid operationId,
            DateTimeOffset startedAt,
            string code,
            string subject,
            string message,
            MigrationExecutionRecorder recorder)
        {
            recorder.RecordState(MigrationExecutionStatus.NotStarted, message);
            return FailureReceipt(package, operationId, startedAt, recorder, MigrationExecutionStatus.NotStarted, false, message, new ExecutionAdmissionFailure
            {
                Code = code,
                Subject = subject,
                Message = message
            });
        }

        public static PublishingPageImportReceipt UnexpectedFailure(
            PublishingPageMigrationPackage package,
            Guid operationId,
            DateTimeOffset startedAt,
            Exception exception,
            MigrationExecutionRecorder recorder)
        {
            return FailureReceipt(
                package,
                operationId,
                startedAt,
                recorder,
                MigrationExecutionStatus.FailedUnexpectedly,
                recorder.Steps.Any(step => step.Outcome == MutationOutcome.Applied || step.Outcome == MutationOutcome.Failed),
                exception.Message,
                null);
        }

        private static PublishingPageImportReceipt FailureReceipt(
            PublishingPageMigrationPackage package,
            Guid operationId,
            DateTimeOffset startedAt,
            MigrationExecutionRecorder recorder,
            MigrationExecutionStatus status,
            bool mutationStarted,
            string message,
            ExecutionAdmissionFailure admissionFailure)
        {
            var frontier = package.Plan.ExecutionFrontier;
            return new PublishingPageImportReceipt
            {
                OperationId = operationId,
                StartedAtUtc = startedAt,
                CompletedAtUtc = DateTimeOffset.UtcNow,
                ApprovedPlanDigest = package.PlanDigest,
                TargetWebUrl = package.Plan.TargetWebUrl,
                TargetPageServerRelativeUrl = package.Plan.TargetPageServerRelativeUrl,
                ExecutionStatus = status,
                PartialExecution = frontier?.IsPartial == true,
                ExecutionFrontier = frontier,
                DeferredIngredientCount = frontier?.Decisions.Count(value => value != null
                    && (value.State == PageIngredientExecutionState.Deferred
                        || value.State == PageIngredientExecutionState.SkippedByDeferredDependency)) ?? 0,
                AuthorizationBlockedIngredientCount = frontier?.Decisions.Count(value => value != null
                    && (value.State == PageIngredientExecutionState.AuthorizationBlocked
                        || value.State == PageIngredientExecutionState.SkippedByAuthorizationDependency)) ?? 0,
                ApprovedLifecycle = package.Plan.TargetLifecycle,
                ExpectedLifecycle = package.Plan.TargetLifecycle,
                AdmissionFailure = admissionFailure,
                MutationStarted = mutationStarted,
                Steps = recorder.Steps,
                StorageVerificationStatus = status == MigrationExecutionStatus.NotStarted
                    ? StorageVerificationStatus.NotRun
                    : StorageVerificationStatus.Failed,
                RuntimeVerificationStatus = RuntimeVerificationStatus.NotRun,
                AcceptanceStatus = MigrationAcceptanceStatus.Rejected,
                Warnings = new List<string> { message }
            };
        }
    }
}
