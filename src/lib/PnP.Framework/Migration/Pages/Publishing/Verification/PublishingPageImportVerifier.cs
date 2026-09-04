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
using PnP.Framework.Migration.Pages.Publishing.Execution;
using PnP.Framework.Migration.Pages.Ingredients;
using PnP.Framework.Migration.Pages.References;
using PnP.Framework.Migration.Pages.Content;
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
            PublishingPageExecutionScope executionScope,
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
            if (!executionScope.PageArtifact)
            {
                return VerifyComponents(
                    package,
                    executionScope,
                    approvedPlanDigest,
                    operationId,
                    startedAt,
                    materializedDependencyCount,
                    topologyReceipt,
                    listReceipts,
                    steps,
                    warnings);
            }
            using (var verificationContext = targetContext.Clone(package.Plan.TargetWebUrl))
            {
                var pages = verificationContext.Web.GetList(
                    package.Plan.TargetProbe?.PagesLibraryServerRelativeUrl
                    ?? PagePath.GetDirectoryName(package.Plan.TargetPageServerRelativeUrl));

                var file = verificationContext.Web.GetFileByServerRelativePath(ResourcePath.FromDecodedUrl(package.Plan.TargetPageServerRelativeUrl));
                var executableTaxonomyActions = executionScope.TaxonomyActions(package).ToArray();
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
                    value => value.CheckOutType,
                    value => value.Properties);
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
                var executableWebPartActions = executionScope.WebPartActions(package);
                var executableWebPartIds = new HashSet<Guid>(
                    executableWebPartActions.Select(value => value.SourceWebPartId));
                var executableWebParts = package.Snapshot.WebParts
                    .Where(value => executableWebPartIds.Contains(value.Id))
                    .ToArray();
                var executionReplacements = PublishingPageExecutionReplacementProjector.Project(
                    package,
                    executionScope);
                var webPartResults = PublishingPageWebPartVerifier.Verify(
                    verificationContext,
                    package.Plan.TargetPageServerRelativeUrl,
                    executableWebParts,
                    package.Snapshot.ListWebPartBindings,
                    executableWebPartActions,
                    listReceipts,
                    executionReplacements,
                    executableWebParts.Length == package.Snapshot.WebParts.Count);
                var persistedDigest = PublishingPageDigest.ComputeSha256(content);
                var expectedExecutionContent = PageTextTransformer.Rewrite(
                    package.Snapshot.PublishingPageContent,
                    executionReplacements);
                var expectedExecutionContentDigest = PublishingPageDigest.ComputeSha256(expectedExecutionContent);
                var receiptWarnings = warnings
                    .Where(value => !string.IsNullOrWhiteSpace(value))
                    .Distinct(StringComparer.Ordinal)
                    .ToList();
                var storageContentEqual = !executionScope.PublishingContent
                    || PublishingPageContentStorageCanonicalizer.AreEquivalent(
                        expectedExecutionContent,
                        content);
                if (!storageContentEqual)
                {
                    receiptWarnings.Add("PublishingPageContent storage bytes differ from the approved digest. Storage verification failed; runtime verification cannot override this mismatch.");
                }
                else if (executionScope.PublishingContent
                    && !string.Equals(persistedDigest, expectedExecutionContentDigest, StringComparison.OrdinalIgnoreCase))
                {
                    receiptWarnings.Add(
                        "SharePoint normalized equivalent HTML character references while persisting PublishingPageContent; canonical authored content matches the approved plan.");
                }

                var expectedContentPresent = executionScope.PublishingContent
                    && !string.IsNullOrWhiteSpace(package.Snapshot.PublishingPageContent);
                var persistedContentPresent = !string.IsNullOrWhiteSpace(content);
                if (expectedContentPresent && !persistedContentPresent)
                {
                    receiptWarnings.Add("Fresh target readback found empty PublishingPageContent even though the approved source snapshot was non-empty.");
                }

                var actualLevel = file.Level.ToString();
                var actualCheckOutType = file.CheckOutType.ToString();
                var effectiveLifecycle = executionScope.Lifecycle
                    ? package.Plan.TargetLifecycle
                    : PublishingPageTargetLifecycle.Draft;
                var lifecycleMatched = effectiveLifecycle == PublishingPageTargetLifecycle.Published
                    ? string.Equals(actualLevel, "Published", StringComparison.OrdinalIgnoreCase)
                    : string.Equals(actualLevel, "Draft", StringComparison.OrdinalIgnoreCase)
                        && string.Equals(actualCheckOutType, "None", StringComparison.OrdinalIgnoreCase);
                if (!lifecycleMatched)
                {
                    receiptWarnings.Add($"Target lifecycle mismatch. Expected {effectiveLifecycle}; actual level is {actualLevel} and checkout state is {actualCheckOutType}.");
                }

                var securityMatched = !executionScope.Security
                    || package.Snapshot.Security.HasUniqueRoleAssignments
                    || !item.HasUniqueRoleAssignments;
                if (!securityMatched)
                {
                    receiptWarnings.Add("Target security mismatch. The source page inherited permissions, but the target page has unique role assignments.");
                }
                var ownershipMatched = string.Equals(
                        Property(file.Properties, PublishingPageTargetOwnership.OriginalIdentifierPropertyName),
                        package.Plan.OriginalIdentifier,
                        StringComparison.Ordinal)
                    && string.Equals(
                        Property(file.Properties, PublishingPageTargetOwnership.SourceSnapshotDigestPropertyName),
                        package.SnapshotDigest,
                        StringComparison.OrdinalIgnoreCase)
                    && string.Equals(
                        Property(file.Properties, PublishingPageTargetOwnership.PlanDigestPropertyName),
                        package.PlanDigest,
                        StringComparison.OrdinalIgnoreCase);
                if (!ownershipMatched)
                {
                    receiptWarnings.Add("Target Page ownership provenance differs from the approved source identity, snapshot digest, or plan digest.");
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
                    && webPartResults.Count == executableWebParts.Length;
                var topologyMatched = TopologyMatched(executionScope, topologyReceipt);
                var listsMatched = ListsMatched(executionScope, listReceipts);
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
                var contentTypeMatched = !executionScope.ContentType
                    || PublishingPageContentTypeIdentity.MatchesSiteContentType(
                        contentTypeId,
                        expectedContentTypeId);
                if (!contentTypeMatched)
                {
                    receiptWarnings.Add($"Target Content Type mismatch. Expected '{expectedContentTypeId ?? "unavailable"}'; actual '{contentTypeId}'.");
                }
                var expectedMaterializedDependencies = executionScope.ReferenceActions(package)
                    .Count(value => value.Disposition == PageReferenceDisposition.MaterializeAtTarget);
                var dependenciesMatched = materializedDependencyCount == expectedMaterializedDependencies;
                if (!dependenciesMatched)
                {
                    receiptWarnings.Add(
                        $"Materialized dependency count differs. Expected {expectedMaterializedDependencies}; observed {materializedDependencyCount}.");
                }
                var readbackPassed = contentTypeMatched
                    && storageContentEqual
                    && webPartsMatched
                    && lifecycleMatched
                    && securityMatched
                    && ownershipMatched
                    && plannedFieldsPassed
                    && taxonomyRelationshipsMatched
                    && topologyMatched
                    && listsMatched
                    && dependenciesMatched;
                AddFrontierWarnings(receiptWarnings, executionScope);
                var runtimeVerificationRequired = executionScope.Runtime
                    && package.Plan.RuntimeVerification.Requirements.Any(item => item.Required);
                var completedIngredientIds = executionScope.ExecutableIngredientIds;
                var ingredientVerification = PublishingPageIngredientVerificationProjector.Project(
                    package,
                    executionScope,
                    new PublishingPageIngredientVerificationEvidence
                    {
                        StructuralMaterializersPassed = true,
                        PageArtifactMatched = ownershipMatched,
                        ContentTypeMatched = contentTypeMatched,
                        PublishingContentMatched = storageContentEqual,
                        SecurityMatched = securityMatched,
                        LifecycleMatched = lifecycleMatched,
                        TaxonomyRelationshipsMatched = taxonomyRelationshipsMatched,
                        TopologyMatched = topologyMatched,
                        DependenciesMatched = dependenciesMatched,
                        RuntimeVerificationRequired = runtimeVerificationRequired,
                        FieldResults = fieldResults,
                        WebPartResults = webPartResults,
                        ListReceipts = listReceipts
                    });
                return new PublishingPageImportReceipt
                {
                    StartedAtUtc = startedAt,
                    CompletedAtUtc = DateTimeOffset.UtcNow,
                    OperationId = operationId,
                    ExecutionStatus = readbackPassed
                        ? SuccessStatus(executionScope)
                        : MigrationExecutionStatus.FailedUnexpectedly,
                    PartialExecution = executionScope.IsPartial,
                    ExecutionFrontier = executionScope.Frontier,
                    CompletedIngredientIds = completedIngredientIds,
                    VerifiedIngredientIds = ingredientVerification.VerifiedIngredientIds,
                    PendingVerificationIngredientIds = ingredientVerification.PendingIngredientIds,
                    FailedVerificationIngredientIds = ingredientVerification.FailedIngredientIds,
                    DeferredIngredientCount = DeferredCount(executionScope),
                    AuthorizationBlockedIngredientCount = AuthorizationBlockedCount(executionScope),
                    MutationStarted = true,
                    Steps = steps,
                    ApprovedPlanDigest = approvedPlanDigest,
                    TargetWebUrl = package.Plan.TargetWebUrl,
                    TargetPageServerRelativeUrl = package.Plan.TargetPageServerRelativeUrl,
                    TargetFileUniqueId = file.UniqueId,
                    TargetListItemId = item.Id,
                    TargetContentTypeId = contentTypeId,
                    TargetVersionLabel = file.UIVersionLabel,
                    ExpectedLifecycle = effectiveLifecycle,
                    ApprovedLifecycle = package.Plan.TargetLifecycle,
                    ActualFileLevel = actualLevel,
                    ActualCheckOutType = actualCheckOutType,
                    ActualModerationStatus = PublishingPageCaptureReader.TryGetInt32(item, "_ModerationStatus"),
                    LifecycleMatched = lifecycleMatched,
                    SecurityMatched = securityMatched,
                    OwnershipMatched = ownershipMatched,
                    PageArtifactMatched = ownershipMatched,
                    LayoutMatched = true,
                    ContentTypeMatched = contentTypeMatched,
                    PageFieldsMatched = plannedFieldsPassed,
                    DependenciesMatched = dependenciesMatched,
                    ApprovedPublishingPageContentSha256 = package.Plan.ExpectedPublishingPageContentSha256,
                    ExpectedPublishingPageContentSha256 = expectedExecutionContentDigest,
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
                            : executionScope.IsPartial
                                ? MigrationAcceptanceStatus.PartiallyAccepted
                                : MigrationAcceptanceStatus.Accepted,
                    Warnings = receiptWarnings.Distinct(StringComparer.Ordinal).OrderBy(value => value, StringComparer.Ordinal).ToList(),
                };
            }
        }

        private static PublishingPageImportReceipt VerifyComponents(
            PublishingPageMigrationPackage package,
            PublishingPageExecutionScope executionScope,
            string approvedPlanDigest,
            Guid operationId,
            DateTimeOffset startedAt,
            int materializedDependencyCount,
            TopologyMaterializationReceipt topologyReceipt,
            IList<ListMaterializationReceipt> listReceipts,
            IList<MigrationMutationReceipt> steps,
            IEnumerable<string> warnings)
        {
            var receiptWarnings = warnings
                .Where(value => !string.IsNullOrWhiteSpace(value))
                .Distinct(StringComparer.Ordinal)
                .ToList();
            var topologyMatched = TopologyMatched(executionScope, topologyReceipt);
            var listsMatched = ListsMatched(executionScope, listReceipts);
            var expectedMaterializedDependencies = executionScope.ReferenceActions(package)
                .Count(value => value.Disposition == PageReferenceDisposition.MaterializeAtTarget);
            var dependenciesMatched = materializedDependencyCount == expectedMaterializedDependencies;
            if (!topologyMatched)
            {
                receiptWarnings.Add("Fresh topology readback did not verify every Site/Web mapping in the execution frontier.");
            }
            if (!listsMatched)
            {
                receiptWarnings.Add("Fresh List readback did not verify every List transaction in the execution frontier.");
            }
            if (!dependenciesMatched)
            {
                receiptWarnings.Add(
                    $"Materialized dependency count differs. Expected {expectedMaterializedDependencies}; observed {materializedDependencyCount}.");
            }
            foreach (var listReceipt in listReceipts.Where(value => !value.FreshReadbackPassed))
            {
                receiptWarnings.AddRange(listReceipt.Diagnostics.Select(value =>
                    "List " + listReceipt.SourceListId.ToString("D") + ": " + value));
            }
            AddFrontierWarnings(receiptWarnings, executionScope);

            var readbackPassed = topologyMatched && listsMatched && dependenciesMatched;
            var completedIngredientIds = executionScope.ExecutableIngredientIds;
            var ingredientVerification = PublishingPageIngredientVerificationProjector.Project(
                package,
                executionScope,
                new PublishingPageIngredientVerificationEvidence
                {
                    StructuralMaterializersPassed = true,
                    PageArtifactMatched = false,
                    ContentTypeMatched = true,
                    PublishingContentMatched = false,
                    SecurityMatched = false,
                    LifecycleMatched = false,
                    TaxonomyRelationshipsMatched = true,
                    TopologyMatched = topologyMatched,
                    DependenciesMatched = dependenciesMatched,
                    RuntimeVerificationRequired = false,
                    ListReceipts = listReceipts
                });
            return new PublishingPageImportReceipt
            {
                StartedAtUtc = startedAt,
                CompletedAtUtc = DateTimeOffset.UtcNow,
                OperationId = operationId,
                ExecutionStatus = readbackPassed
                    ? SuccessStatus(executionScope)
                    : MigrationExecutionStatus.FailedUnexpectedly,
                PartialExecution = executionScope.IsPartial,
                ExecutionFrontier = executionScope.Frontier,
                CompletedIngredientIds = completedIngredientIds,
                VerifiedIngredientIds = ingredientVerification.VerifiedIngredientIds,
                PendingVerificationIngredientIds = ingredientVerification.PendingIngredientIds,
                FailedVerificationIngredientIds = ingredientVerification.FailedIngredientIds,
                DeferredIngredientCount = DeferredCount(executionScope),
                AuthorizationBlockedIngredientCount = AuthorizationBlockedCount(executionScope),
                MutationStarted = true,
                Steps = steps,
                ApprovedPlanDigest = approvedPlanDigest,
                TargetWebUrl = package.Plan.TargetWebUrl,
                TargetPageServerRelativeUrl = package.Plan.TargetPageServerRelativeUrl,
                ApprovedLifecycle = package.Plan.TargetLifecycle,
                ExpectedLifecycle = PublishingPageTargetLifecycle.Draft,
                LifecycleMatched = true,
                SecurityMatched = true,
                OwnershipMatched = true,
                PageArtifactMatched = true,
                LayoutMatched = true,
                ContentTypeMatched = true,
                PageFieldsMatched = true,
                DependenciesMatched = dependenciesMatched,
                StorageContentEqual = true,
                WebPartsMatched = true,
                MaterializedDependencyCount = materializedDependencyCount,
                TopologyMaterialization = topologyReceipt,
                TopologyMatched = topologyMatched,
                ListMaterializations = listReceipts,
                ListsMatched = listsMatched,
                TaxonomyRelationshipsMatched = true,
                FreshReadbackPassed = readbackPassed,
                StorageVerificationStatus = readbackPassed
                    ? StorageVerificationStatus.Passed
                    : StorageVerificationStatus.Failed,
                RuntimeVerificationStatus = RuntimeVerificationStatus.NotRequired,
                AcceptanceStatus = !readbackPassed
                    ? MigrationAcceptanceStatus.Rejected
                    : executionScope.IsPartial
                        ? MigrationAcceptanceStatus.PartiallyAccepted
                        : MigrationAcceptanceStatus.Accepted,
                Warnings = receiptWarnings
                    .Distinct(StringComparer.Ordinal)
                    .OrderBy(value => value, StringComparer.Ordinal)
                    .ToList()
            };
        }

        private static bool TopologyMatched(
            PublishingPageExecutionScope executionScope,
            TopologyMaterializationReceipt receipt)
        {
            var expected = executionScope.TopologyPlan?.SiteCollections
                .SelectMany(value => value.Webs)
                .Count() ?? 0;
            return receipt != null
                && receipt.FreshReadbackPassed
                && receipt.Webs.Count == expected;
        }

        private static bool ListsMatched(
            PublishingPageExecutionScope executionScope,
            IList<ListMaterializationReceipt> receipts)
        {
            var expected = executionScope.ListScope?.Lists.Count(value => value.HasListScopedWork) ?? 0;
            return receipts != null
                && receipts.Count == expected
                && receipts.All(value => value.FreshReadbackPassed);
        }

        private static MigrationExecutionStatus SuccessStatus(PublishingPageExecutionScope executionScope)
        {
            return executionScope.IsPartial
                ? MigrationExecutionStatus.PartiallySucceeded
                : MigrationExecutionStatus.Succeeded;
        }

        private static int DeferredCount(PublishingPageExecutionScope executionScope)
        {
            return executionScope.Frontier.Decisions.Count(value => value != null
                && (value.State == PageIngredientExecutionState.Deferred
                    || value.State == PageIngredientExecutionState.SkippedByDeferredDependency));
        }

        private static int AuthorizationBlockedCount(PublishingPageExecutionScope executionScope)
        {
            return executionScope.Frontier.Decisions.Count(value => value != null
                && (value.State == PageIngredientExecutionState.AuthorizationBlocked
                    || value.State == PageIngredientExecutionState.SkippedByAuthorizationDependency));
        }

        private static void AddFrontierWarnings(
            ICollection<string> warnings,
            PublishingPageExecutionScope executionScope)
        {
            var deferred = DeferredCount(executionScope);
            var authorizationBlocked = AuthorizationBlockedCount(executionScope);
            if (deferred > 0)
            {
                warnings.Add(
                    deferred + " ingredient(s) remain deferred or were skipped because a required dependency is deferred; their snapshot evidence remains sealed for a later run.");
            }
            if (authorizationBlocked > 0)
            {
                warnings.Add(
                    authorizationBlocked + " ingredient(s) are blocked by retained literal HTTP 401/403 evidence or a required dependency on that evidence; independent ingredients continued.");
            }
        }

        private static string Property(PropertyValues values, string name)
        {
            object value;
            return values != null && values.FieldValues.TryGetValue(name, out value) ? Convert.ToString(value) : null;
        }
    }
}
