using Microsoft.SharePoint.Client;
using PnP.Framework.Migration.Pages.Capture;
using PnP.Framework.Migration.Pages.Fields;
using PnP.Framework.Migration.Pages.Lifecycle;
using PnP.Framework.Migration.Pages.Publishing.Capture;
using PnP.Framework.Migration.Pages.Publishing.Lifecycle;
using PnP.Framework.Migration.Pages.Publishing.Packaging;
using PnP.Framework.Migration.Execution;
using PnP.Framework.Migration.Verification;
using System;
using System.Collections.Generic;
using System.Linq;

namespace PnP.Framework.Migration.Pages.Publishing.Verification
{
    internal static class PublishingPageImportVerifier
    {
        public static PublishingPageImportReceipt Verify(
            ClientContext targetContext,
            PublishingPageMigrationPackage package,
            string approvedPlanDigest,
            Guid operationId,
            DateTimeOffset startedAt,
            int materializedDependencyCount,
            IList<PageFieldImportResult> fieldResults,
            IList<MigrationMutationReceipt> steps,
            IEnumerable<string> warnings,
            Func<string, bool> isExpectedContentType)
        {
            using (var verificationContext = targetContext.Clone(package.Plan.TargetWebUrl))
            {
                var pages = verificationContext.Web.GetPagesLibrary();
                if (pages == null)
                {
                    throw new InvalidOperationException("Fresh target readback could not find the publishing Pages library.");
                }

                var file = verificationContext.Web.GetFileByServerRelativePath(ResourcePath.FromDecodedUrl(package.Plan.TargetPageServerRelativeUrl));
                var items = pages.GetItems(new CamlQuery
                {
                    ViewXml = $@"<View Scope='RecursiveAll'>
  <Query>
    <Where>
      <Eq>
        <FieldRef Name='FileRef' />
        <Value Type='Text'>{System.Security.SecurityElement.Escape(package.Plan.TargetPageServerRelativeUrl)}</Value>
      </Eq>
    </Where>
  </Query>
  <ViewFields>
    <FieldRef Name='ID' />
    <FieldRef Name='ContentTypeId' />
    <FieldRef Name='PublishingPageContent' />
    <FieldRef Name='_ModerationStatus' />
  </ViewFields>
  <RowLimit>1</RowLimit>
</View>"
                });
                verificationContext.Load(file,
                    value => value.Exists,
                    value => value.UniqueId,
                    value => value.UIVersionLabel,
                    value => value.Level,
                    value => value.CheckOutType);
                verificationContext.Load(items);
                verificationContext.ExecuteQueryRetry();
                if (!file.Exists)
                {
                    throw new InvalidOperationException("Fresh target readback could not find the imported page.");
                }

                var item = items.SingleOrDefault();
                if (item == null)
                {
                    throw new InvalidOperationException("Fresh target readback could not find the imported page list item.");
                }

                var content = PublishingPageCaptureReader.GetFieldString(item, "PublishingPageContent") ?? string.Empty;
                var contentTypeId = PublishingPageCaptureReader.GetFieldString(item, "ContentTypeId") ?? string.Empty;
                var webPartResults = PublishingPageWebPartVerifier.Verify(
                    verificationContext,
                    package.Plan.TargetPageServerRelativeUrl,
                    package.Snapshot.WebParts,
                    package.Plan.Replacements);
                var persistedDigest = PublishingPageDigest.ComputeSha256(content);
                var receiptWarnings = warnings
                    .Where(value => !string.IsNullOrWhiteSpace(value))
                    .Distinct(StringComparer.Ordinal)
                    .ToList();
                var storageContentEqual = string.Equals(
                    persistedDigest,
                    package.Plan.ExpectedPublishingPageContentSha256,
                    StringComparison.OrdinalIgnoreCase);
                if (!storageContentEqual)
                {
                    receiptWarnings.Add("PublishingPageContent storage bytes differ from the approved digest. Storage verification failed; runtime verification cannot override this mismatch.");
                }

                var expectedContentPresent = !string.IsNullOrWhiteSpace(package.Snapshot.PublishingPageContent);
                var persistedContentPresent = !string.IsNullOrWhiteSpace(content);
                if (expectedContentPresent && !persistedContentPresent)
                {
                    receiptWarnings.Add("Fresh target readback found empty PublishingPageContent even though the approved source snapshot was non-empty.");
                }

                var actualLevel = file.Level.ToString();
                var actualCheckOutType = file.CheckOutType.ToString();
                var lifecycleMatched = package.Plan.TargetLifecycle == PublishingPageTargetLifecycle.Published
                    ? string.Equals(actualLevel, "Published", StringComparison.OrdinalIgnoreCase)
                    : string.Equals(actualLevel, "Draft", StringComparison.OrdinalIgnoreCase)
                        && string.Equals(actualCheckOutType, "None", StringComparison.OrdinalIgnoreCase);
                if (!lifecycleMatched)
                {
                    receiptWarnings.Add($"Target lifecycle mismatch. Expected {package.Plan.TargetLifecycle}; actual level is {actualLevel} and checkout state is {actualCheckOutType}.");
                }

                var plannedFieldsPassed = fieldResults.All(result => !result.Attempted || result.Succeeded);
                var webPartsMatched = webPartResults.All(result => result.Passed)
                    && webPartResults.Count == package.Snapshot.WebParts.Count;
                var readbackPassed = isExpectedContentType(contentTypeId)
                    && storageContentEqual
                    && webPartsMatched
                    && lifecycleMatched
                    && plannedFieldsPassed;
                var runtimeVerificationRequired = package.Plan.RuntimeVerification.Requirements.Any(item => item.Required);
                return new PublishingPageImportReceipt
                {
                    StartedAtUtc = startedAt,
                    CompletedAtUtc = DateTimeOffset.UtcNow,
                    OperationId = operationId,
                    ExecutionStatus = readbackPassed
                        ? MigrationExecutionStatus.Succeeded
                        : MigrationExecutionStatus.FailedUnexpectedly,
                    MutationStarted = true,
                    Steps = steps,
                    ApprovedPlanDigest = approvedPlanDigest,
                    TargetWebUrl = package.Plan.TargetWebUrl,
                    TargetPageServerRelativeUrl = package.Plan.TargetPageServerRelativeUrl,
                    TargetFileUniqueId = file.UniqueId,
                    TargetListItemId = item.Id,
                    TargetContentTypeId = contentTypeId,
                    TargetVersionLabel = file.UIVersionLabel,
                    ExpectedLifecycle = package.Plan.TargetLifecycle,
                    ActualFileLevel = actualLevel,
                    ActualCheckOutType = actualCheckOutType,
                    ActualModerationStatus = PublishingPageCaptureReader.TryGetInt32(item, "_ModerationStatus"),
                    LifecycleMatched = lifecycleMatched,
                    ExpectedPublishingPageContentSha256 = package.Plan.ExpectedPublishingPageContentSha256,
                    PersistedPublishingPageContentSha256 = persistedDigest,
                    StorageContentEqual = storageContentEqual,
                    ImportedWebPartCount = webPartResults.Count(result => result.TargetWebPartId.HasValue),
                    WebPartsMatched = webPartsMatched,
                    WebPartResults = webPartResults,
                    MaterializedDependencyCount = materializedDependencyCount,
                    FieldResults = fieldResults,
                    FreshReadbackPassed = readbackPassed,
                    StorageVerificationStatus = readbackPassed
                        ? StorageVerificationStatus.Passed
                        : StorageVerificationStatus.Failed,
                    RuntimeVerificationStatus = runtimeVerificationRequired
                        ? RuntimeVerificationStatus.Pending
                        : RuntimeVerificationStatus.NotRequired,
                    AcceptanceStatus = !readbackPassed
                        ? MigrationAcceptanceStatus.Rejected
                        : runtimeVerificationRequired
                            ? MigrationAcceptanceStatus.Pending
                            : MigrationAcceptanceStatus.Accepted,
                    Warnings = receiptWarnings.Distinct(StringComparer.Ordinal).OrderBy(value => value, StringComparer.Ordinal).ToList(),
                };
            }
        }
    }
}
