using Microsoft.SharePoint.Client;
using PnP.Framework.Migration.Pages.Capture;
using PnP.Framework.Migration.Pages.Fields;
using PnP.Framework.Migration.Pages.Lifecycle;
using PnP.Framework.Migration.Pages.Publishing.Capture;
using PnP.Framework.Migration.Pages.Publishing.Lifecycle;
using PnP.Framework.Migration.Pages.Publishing.Packaging;
using PnP.Framework.Migration.Execution;
using PnP.Framework.Migration.Verification;
using PnP.Framework.Migration.Lists.Planning;
using PnP.Framework.Migration.Topology;
using PnP.Framework.Migration.Pages.Fields.Taxonomy;
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
            TopologyMaterializationReceipt topologyReceipt,
            IList<ListMaterializationReceipt> listReceipts,
            IList<PageFieldImportResult> fieldResults,
            IList<MigrationMutationReceipt> steps,
            IEnumerable<string> warnings,
            string expectedContentTypeId)
        {
            using (var verificationContext = targetContext.Clone(package.Plan.TargetWebUrl))
            {
                var pages = verificationContext.Web.GetList(
                    PagePath.GetDirectoryName(package.Plan.TargetPageServerRelativeUrl));

                var file = verificationContext.Web.GetFileByServerRelativePath(ResourcePath.FromDecodedUrl(package.Plan.TargetPageServerRelativeUrl));
                var executableTaxonomyActions = package.Plan.TaxonomyRelationshipActions
                    .Where(value => value.IsExecutable)
                    .ToArray();
                var taxonomyViewFields = string.Join(
                    Environment.NewLine,
                    executableTaxonomyActions
                        .Select(value => value.SourceFieldInternalName)
                        .Where(value => !string.IsNullOrWhiteSpace(value))
                        .Distinct(StringComparer.OrdinalIgnoreCase)
                        .OrderBy(value => value, StringComparer.Ordinal)
                        .Select(value => "    <FieldRef Name='" + System.Security.SecurityElement.Escape(value) + "' />"));
                var taxCatchAllViewField = executableTaxonomyActions.Length > 0
                    ? "    <FieldRef Name='TaxCatchAll' />"
                    : string.Empty;
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
{taxonomyViewFields}
{taxCatchAllViewField}
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

                verificationContext.Load(item, value => value.HasUniqueRoleAssignments);
                verificationContext.ExecuteQueryRetry();

                var content = PublishingPageCaptureReader.GetFieldString(item, "PublishingPageContent") ?? string.Empty;
                var contentTypeId = PublishingPageCaptureReader.GetFieldString(item, "ContentTypeId") ?? string.Empty;
                var webPartResults = PublishingPageWebPartVerifier.Verify(
                    verificationContext,
                    package.Plan.TargetPageServerRelativeUrl,
                    package.Snapshot.WebParts,
                    package.Snapshot.ListWebPartBindings,
                    package.Plan.WebPartActions,
                    listReceipts,
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

                var securityMatched = package.Snapshot.Security.HasUniqueRoleAssignments
                    || !item.HasUniqueRoleAssignments;
                if (!securityMatched)
                {
                    receiptWarnings.Add("Target security mismatch. The source page inherited permissions, but the target page has unique role assignments.");
                }

                var plannedFieldsPassed = fieldResults.All(result => !result.Attempted || result.Succeeded);
                var taxonomyRelationshipResults = PageTaxonomyRelationshipVerifier.Verify(
                    verificationContext,
                    pages,
                    item,
                    package.Snapshot.Fields,
                    executableTaxonomyActions,
                    fieldResults);
                var taxonomyRelationshipsMatched = taxonomyRelationshipResults.All(value => value.Passed)
                    && taxonomyRelationshipResults.Count == executableTaxonomyActions.Length;
                if (!taxonomyRelationshipsMatched)
                {
                    receiptWarnings.Add("Fresh taxonomy readback did not reproduce every sealed relationship exactly.");
                    receiptWarnings.AddRange(taxonomyRelationshipResults
                        .Where(value => !value.Passed)
                        .Select(value => value.SourceFieldInternalName + ":" + value.SourceTermId.ToString("D") + ": " + value.Message));
                }
                var webPartsMatched = webPartResults.All(result => result.Passed)
                    && webPartResults.Count == package.Snapshot.WebParts.Count;
                var topologyMatched = topologyReceipt != null
                    && topologyReceipt.FreshReadbackPassed
                    && (package.Plan.Topology == null
                        || topologyReceipt.Webs.Count == package.Plan.Topology.SiteCollections.SelectMany(value => value.Webs).Count());
                var listsMatched = listReceipts.Count == package.Snapshot.ListDependencies.Count
                    && listReceipts.All(value => value.FreshReadbackPassed);
                if (!topologyMatched)
                {
                    receiptWarnings.Add("Fresh topology readback did not verify every approved Site/Web mapping.");
                }
                foreach (var listReceipt in listReceipts.Where(value => !value.FreshReadbackPassed))
                {
                    receiptWarnings.AddRange(listReceipt.Diagnostics.Select(value => "List " + listReceipt.SourceListId.ToString("D") + ": " + value));
                }
                if (!listsMatched)
                {
                    receiptWarnings.Add("Fresh List readback did not verify every captured List dependency.");
                }
                var contentTypeMatched = !string.IsNullOrWhiteSpace(expectedContentTypeId)
                    && string.Equals(contentTypeId, expectedContentTypeId, StringComparison.OrdinalIgnoreCase);
                if (!contentTypeMatched)
                {
                    receiptWarnings.Add($"Target Content Type mismatch. Expected '{expectedContentTypeId ?? "unavailable"}'; actual '{contentTypeId}'.");
                }
                var readbackPassed = contentTypeMatched
                    && storageContentEqual
                    && webPartsMatched
                    && lifecycleMatched
                    && securityMatched
                    && plannedFieldsPassed
                    && taxonomyRelationshipsMatched
                    && topologyMatched
                    && listsMatched;
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
                    SecurityMatched = securityMatched,
                    ExpectedPublishingPageContentSha256 = package.Plan.ExpectedPublishingPageContentSha256,
                    PersistedPublishingPageContentSha256 = persistedDigest,
                    StorageContentEqual = storageContentEqual,
                    ImportedWebPartCount = webPartResults.Count(result => result.TargetWebPartId.HasValue),
                    WebPartsMatched = webPartsMatched,
                    WebPartResults = webPartResults,
                    MaterializedDependencyCount = materializedDependencyCount,
                    TopologyMaterialization = topologyReceipt,
                    TopologyMatched = topologyMatched,
                    ListMaterializations = listReceipts,
                    ListsMatched = listsMatched,
                    FieldResults = fieldResults,
                    TaxonomyRelationshipsMatched = taxonomyRelationshipsMatched,
                    TaxonomyRelationshipResults = taxonomyRelationshipResults,
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
