using Microsoft.SharePoint.Client;
using PnP.Framework.Migration.Execution;
using PnP.Framework.Migration.Pages.Publishing.Execution;
using PnP.Framework.Migration.Pages.Publishing.Packaging;
using PnP.Framework.Migration.Packaging;
using System;
using PnP.Framework.Migration.Pages.Publishing.Profiles;

namespace PnP.Framework.Migration.Pages.Publishing.EnterpriseWiki
{
    public sealed class EnterpriseWikiMigrationImporter
    {
        public PublishingPageImportReceipt Import(
            ClientContext targetContext,
            PublishingPageMigrationPackage package,
            string approvedPlanDigest)
        {
            return Import(targetContext, package, approvedPlanDigest, null, null);
        }

        public PublishingPageImportReceipt Import(
            ClientContext targetContext,
            PublishingPageMigrationPackage package,
            string approvedPlanDigest,
            IMigrationExecutionJournal journal)
        {
            return Import(targetContext, package, approvedPlanDigest, journal, null);
        }

        public PublishingPageImportReceipt Import(
            ClientContext targetContext,
            PublishingPageMigrationPackage package,
            string approvedPlanDigest,
            IMigrationExecutionJournal journal,
            IMigrationArtifactStore artifactStore)
        {
            if (targetContext == null)
            {
                throw new ArgumentNullException(nameof(targetContext));
            }

            PublishingPagePackageValidator.ValidateMigration(package, artifactStore);
            var executionScope = PublishingPageExecutionScope.Create(package);
            PublishingPageImportPlanValidator.Validate(package, EnterpriseWikiV1WorkflowPolicy.Instance, executionScope);
            var operationId = Guid.NewGuid();
            var startedAt = DateTimeOffset.UtcNow;
            var recorder = new MigrationExecutionRecorder(operationId, package.PlanDigest, journal);
            var admissionFailure = PublishingPageImportAdmission.TryAdmit(
                targetContext,
                package,
                executionScope,
                approvedPlanDigest,
                operationId,
                startedAt,
                recorder);
            if (admissionFailure != null)
            {
                return admissionFailure;
            }

            recorder.RecordState(MigrationExecutionStatus.Running, "Target admission passed. Mutation execution is starting.");
            try
            {
                return PublishingPageMutationExecutor.Execute(
                    targetContext,
                    package,
                    executionScope,
                    approvedPlanDigest,
                    operationId,
                    startedAt,
                    recorder,
                    artifactStore,
                    package.Plan.TargetProbe?.PageContentTypeId);
            }
            catch (Exception exception)
            {
                recorder.RecordState(MigrationExecutionStatus.FailedUnexpectedly, exception.Message);
                return PublishingPageImportReceiptFactory.UnexpectedFailure(
                    package,
                    operationId,
                    startedAt,
                    exception,
                    recorder);
            }
        }
    }
}
