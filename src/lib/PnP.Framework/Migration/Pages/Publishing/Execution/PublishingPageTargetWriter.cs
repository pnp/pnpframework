using Microsoft.SharePoint.Client;
using PnP.Framework.Entities;
using PnP.Framework.Migration.Execution;
using PnP.Framework.Migration.Pages.Content;
using PnP.Framework.Migration.Pages.Fields;
using PnP.Framework.Migration.Pages.Publishing.Packaging;
using PnP.Framework.Migration.Lists.Planning;
using PnP.Framework.Migration.Pages.ClassicWebParts.Bindings;
using System;
using System.Collections.Generic;
using System.Linq;

namespace PnP.Framework.Migration.Pages.Publishing.Execution
{
    internal static class PublishingPageTargetWriter
    {
        public static PublishingPageWriteResult Write(
            ClientContext targetContext,
            PublishingPageMigrationPackage package,
            PublishingPageExecutionScope executionScope,
            IDictionary<Guid, ListMaterializationReceipt> listReceipts,
            MigrationExecutionRecorder recorder,
            ICollection<string> warnings)
        {
            var targetWeb = targetContext.Web;
            var targetLocation = PublishingPageTargetLocationMaterializer.Ensure(
                targetContext,
                package,
                recorder);
            var pages = targetLocation.PagesLibrary;
            var targetFile = targetWeb.GetFileByServerRelativePath(ResourcePath.FromDecodedUrl(package.Plan.TargetPageServerRelativeUrl));
            targetContext.Load(targetFile, file => file.Exists, file => file.Properties);
            var resumeOwnedPage = false;
            try
            {
                targetContext.ExecuteQueryRetry();
                resumeOwnedPage = targetFile.Exists;
            }
            catch (ServerException exception) when (IsMissing(exception))
            {
                // Some SharePoint farms throw for a missing file instead of
                // returning Exists=false. This is the same compatibility shape
                // handled by the target inspector and remains a create-missing
                // observation, not an unexpected failure.
            }
            if (resumeOwnedPage)
            {
                if (!Verification.PublishingPageTargetOwnership.MatchesApprovedPlan(
                    targetFile.Properties.FieldValues,
                    package.Plan.OriginalIdentifier,
                    package.SnapshotDigest,
                    package.PlanDigest))
                {
                    throw new InvalidOperationException(
                        $"The approved exact page path is occupied by a target that is not owned by this sealed plan: '{package.Plan.TargetPageServerRelativeUrl}'.");
                }
                recorder.RecordAlreadySatisfied(
                    "page.create",
                    $"Resume the exact migration-owned publishing page '{package.Plan.TargetPageServerRelativeUrl}' under the same sealed plan; mutable ingredients will be verified in place rather than replayed.");
            }
            else
            {
                recorder.Execute(
                    "page.create",
                    $"Create publishing page '{package.Plan.TargetPageServerRelativeUrl}'.",
                    () => targetWeb.AddPublishingPage(
                        targetLocation.FileName,
                        package.Plan.PageLayoutName,
                        package.Snapshot.Source.Title,
                        false,
                        targetLocation.TargetFolder));
            }

            targetFile = targetWeb.GetFileByServerRelativePath(ResourcePath.FromDecodedUrl(package.Plan.TargetPageServerRelativeUrl));
            var targetItem = targetFile.ListItemAllFields;
            targetContext.Load(targetFile, file => file.Exists, file => file.CheckOutType, file => file.Properties);
            targetContext.Load(targetItem, item => item.Id);
            targetContext.ExecuteQueryRetry();
            if (!targetFile.Exists)
            {
                throw new InvalidOperationException(
                    $"SharePoint did not create the publishing page at the approved exact path '{package.Plan.TargetPageServerRelativeUrl}'.");
            }
            if (resumeOwnedPage)
            {
                recorder.RecordAlreadySatisfied(
                    "page.checkout",
                    "The exact migration-owned page will be verified in place without opening another edit transaction.");
                recorder.RecordAlreadySatisfied(
                    "page.content",
                    "PublishingPageContent on the exact migration-owned page will be checked by fresh storage readback.");
                recorder.RecordAlreadySatisfied(
                    "page.fields",
                    "Approved page field actions on the resumed page will be checked by fresh storage readback rather than replayed.");
                recorder.RecordAlreadySatisfied(
                    "page.webparts",
                    "Approved shared Web Part actions on the resumed page will be checked by fresh storage readback rather than replayed.");
                recorder.RecordAlreadySatisfied(
                    "page.ownership",
                    "The exact target page already carries matching source identity, snapshot digest, and plan digest provenance.");
                recorder.RecordAlreadySatisfied(
                    "page.security",
                    "Fresh readback will verify the resumed page security policy.");
                return new PublishingPageWriteResult
                {
                    PagesLibrary = pages,
                    TargetFile = targetFile,
                    TargetItem = targetItem,
                    ResumedExistingOwnedPage = true,
                    FieldResults = new List<PageFieldImportResult>()
                };
            }
            EnsureCheckout(targetContext, pages, targetFile, recorder);
            WriteContent(targetContext, targetItem, package, executionScope, recorder);
            var fieldResults = WriteFields(targetContext, targetItem, package, executionScope, recorder, warnings);
            WriteWebParts(targetWeb, package, executionScope, listReceipts, recorder);
            WriteOwnership(targetContext, targetFile, package, recorder);
            recorder.RecordAlreadySatisfied(
                "page.security",
                executionScope.Security
                    ? "The newly created page inherits target Pages-library permissions as required by the admitted security transaction."
                    : "The Page security transaction is outside the admitted execution frontier.");
            return new PublishingPageWriteResult
            {
                PagesLibrary = pages,
                TargetFile = targetFile,
                TargetItem = targetItem,
                FieldResults = fieldResults
            };
        }

        private static bool IsMissing(ServerException exception)
        {
            return string.Equals(exception.ServerErrorTypeName, "System.IO.FileNotFoundException", StringComparison.Ordinal)
                || exception.ServerErrorCode == -2147024894;
        }

        private static void WriteOwnership(
            ClientContext context,
            File targetFile,
            PublishingPageMigrationPackage package,
            MigrationExecutionRecorder recorder)
        {
            recorder.Execute("page.ownership", "Write Page source identity and approved digest provenance.", () =>
            {
                targetFile.Properties[Verification.PublishingPageTargetOwnership.OriginalIdentifierPropertyName] = package.Plan.OriginalIdentifier;
                targetFile.Properties[Verification.PublishingPageTargetOwnership.SourceSnapshotDigestPropertyName] = package.SnapshotDigest;
                targetFile.Properties[Verification.PublishingPageTargetOwnership.PlanDigestPropertyName] = package.PlanDigest;
                targetFile.Update();
                context.ExecuteQueryRetry();
            });
        }

        private static void EnsureCheckout(
            ClientContext context,
            List pages,
            File targetFile,
            MigrationExecutionRecorder recorder)
        {
            if (pages.ForceCheckout && targetFile.CheckOutType == CheckOutType.None)
            {
                recorder.Execute("page.checkout", "Check out the newly created publishing page before field updates.", () =>
                {
                    targetFile.CheckOut();
                    context.ExecuteQueryRetry();
                });
                return;
            }

            recorder.RecordAlreadySatisfied("page.checkout", "The target page does not require an explicit checkout.");
        }

        private static void WriteContent(
            ClientContext context,
            ListItem targetItem,
            PublishingPageMigrationPackage package,
            PublishingPageExecutionScope executionScope,
            MigrationExecutionRecorder recorder)
        {
            if (!executionScope.PublishingContent)
            {
                recorder.RecordAlreadySatisfied(
                    "page.content",
                    "PublishingPageContent is outside the admitted execution frontier; the page shell remains without migrated body content.");
                return;
            }
            var replacements = PublishingPageExecutionReplacementProjector.Project(package, executionScope);
            var rewrittenContent = PageTextTransformer.Rewrite(package.Snapshot.PublishingPageContent, replacements);
            recorder.Execute("page.content", "Write the approved title and PublishingPageContent.", () =>
            {
                targetItem["Title"] = package.Snapshot.Source.Title;
                targetItem["PublishingPageContent"] = rewrittenContent;
                targetItem.Update();
                context.ExecuteQueryRetry();
            });
        }

        private static IList<PageFieldImportResult> WriteFields(
            ClientContext context,
            ListItem targetItem,
            PublishingPageMigrationPackage package,
            PublishingPageExecutionScope executionScope,
            MigrationExecutionRecorder recorder,
            ICollection<string> warnings)
        {
            var fieldActions = executionScope.PageFieldActions(package);
            var taxonomyActions = executionScope.TaxonomyActions(package);
            var replacements = PublishingPageExecutionReplacementProjector.Project(package, executionScope);
            if (fieldActions.Count == 0)
            {
                recorder.RecordAlreadySatisfied(
                    "page.fields",
                    "No Page field-value transaction is present in the admitted execution frontier.");
                return new List<PageFieldImportResult>();
            }
            return recorder.Execute(
                "page.fields",
                "Apply approved page field actions.",
                () => PageFieldWriter.Apply(
                    context,
                    targetItem,
                    package.Snapshot.Fields,
                    fieldActions,
                    taxonomyActions,
                    replacements,
                    warnings),
                results => results.Any(result => result.Attempted)
                    ? results.Any(result => result.Attempted && !result.Succeeded) ? MutationOutcome.Failed : MutationOutcome.Applied
                    : MutationOutcome.AlreadySatisfied,
                results => $"Attempted {results.Count(result => result.Attempted)} approved field write(s); {results.Count(result => result.Attempted && !result.Succeeded)} failed.");
        }

        private static void WriteWebParts(
            Web targetWeb,
            PublishingPageMigrationPackage package,
            PublishingPageExecutionScope executionScope,
            IDictionary<Guid, ListMaterializationReceipt> listReceipts,
            MigrationExecutionRecorder recorder)
        {
            var selectedActions = executionScope.WebPartActions(package);
            if (selectedActions.Count == 0)
            {
                recorder.RecordAlreadySatisfied(
                    "page.webparts",
                    "No shared Web Part transaction is present in the admitted execution frontier.");
                return;
            }
            var replacements = PublishingPageExecutionReplacementProjector.Project(package, executionScope);

            var actions = selectedActions.ToDictionary(value => value.SourceWebPartId);
            var bindings = package.Snapshot.ListWebPartBindings.ToDictionary(value => value.SourceWebPartId);
            foreach (var webPart in package.Snapshot.WebParts.Where(value => actions.ContainsKey(value.Id)))
            {
                var captured = webPart;
                ClassicWebPartAction action;
                if (!actions.TryGetValue(captured.Id, out action))
                {
                    throw new InvalidOperationException("The captured Web Part has no approved replay action: " + captured.Id.ToString("D") + ".");
                }
                ClassicListWebPartBindingSnapshot binding;
                bindings.TryGetValue(captured.Id, out binding);
                ListMaterializationReceipt listReceipt = null;
                if (binding != null && (listReceipts == null || !listReceipts.TryGetValue(binding.SourceListId, out listReceipt)))
                {
                    throw new InvalidOperationException("The list-bound Web Part has no materialized target List receipt: " + captured.Id.ToString("D") + ".");
                }
                var replayXml = ClassicWebPartReplayComposer.Compose(
                    captured,
                    action,
                    binding,
                    listReceipt,
                    package.Plan.TargetPageServerRelativeUrl,
                    replacements);
                recorder.Execute(
                    $"page.webpart.{captured.Id:N}",
                    $"Import Web Part '{captured.Title}' into zone '{captured.ZoneId}' at index {captured.ZoneIndex}.",
                    () => targetWeb.AddWebPartToWebPartPage(package.Plan.TargetPageServerRelativeUrl, new WebPartEntity
                    {
                        WebPartIndex = captured.ZoneIndex,
                        WebPartTitle = captured.Title,
                        WebPartZone = captured.ZoneId,
                        WebPartXml = replayXml
                    }));
            }
        }
    }
}
