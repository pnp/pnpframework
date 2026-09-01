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
            IDictionary<Guid, ListMaterializationReceipt> listReceipts,
            MigrationExecutionRecorder recorder,
            ICollection<string> warnings)
        {
            var targetWeb = targetContext.Web;
            var pages = GetPagesLibrary(targetContext, package);
            var targetFileName = PagePath.GetFileName(package.Plan.TargetPageServerRelativeUrl);
            recorder.Execute(
                "page.create",
                $"Create publishing page '{package.Plan.TargetPageServerRelativeUrl}'.",
                () => targetWeb.AddPublishingPage(targetFileName, package.Plan.PageLayoutName, package.Snapshot.Source.Title, false));

            var targetFile = targetWeb.GetFileByServerRelativePath(ResourcePath.FromDecodedUrl(package.Plan.TargetPageServerRelativeUrl));
            var targetItem = targetFile.ListItemAllFields;
            targetContext.Load(targetFile, file => file.Exists, file => file.CheckOutType);
            targetContext.Load(targetItem, item => item.Id);
            targetContext.ExecuteQueryRetry();
            EnsureCheckout(targetContext, pages, targetFile, recorder);
            WriteContent(targetContext, targetItem, package, recorder);
            var fieldResults = WriteFields(targetContext, targetItem, package, recorder, warnings);
            WriteWebParts(targetWeb, package, listReceipts, recorder);
            return new PublishingPageWriteResult
            {
                PagesLibrary = pages,
                TargetFile = targetFile,
                TargetItem = targetItem,
                FieldResults = fieldResults
            };
        }

        private static List GetPagesLibrary(ClientContext context, PublishingPageMigrationPackage package)
        {
            var pages = context.Web.GetPagesLibrary();
            if (pages == null)
            {
                throw new InvalidOperationException("The target publishing Pages library is unavailable.");
            }

            context.Load(pages, list => list.EnableModeration, list => list.ForceCheckout);
            context.Load(pages.RootFolder, folder => folder.ServerRelativeUrl);
            context.ExecuteQueryRetry();
            var targetDirectory = PagePath.GetDirectoryName(package.Plan.TargetPageServerRelativeUrl);
            if (!string.Equals(targetDirectory, pages.RootFolder.ServerRelativeUrl, StringComparison.OrdinalIgnoreCase))
            {
                throw new NotSupportedException("The publishing-page importer supports pages in the root of the target Pages library only.");
            }

            return pages;
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
            MigrationExecutionRecorder recorder)
        {
            var rewrittenContent = PageTextTransformer.Rewrite(package.Snapshot.PublishingPageContent, package.Plan.Replacements);
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
            MigrationExecutionRecorder recorder,
            ICollection<string> warnings)
        {
            return recorder.Execute(
                "page.fields",
                "Apply approved page field actions.",
                () => PageFieldWriter.Apply(context, targetItem, package.Snapshot.Fields, package.Plan.FieldActions, package.Plan.Replacements, warnings),
                results => results.Any(result => result.Attempted)
                    ? results.Any(result => result.Attempted && !result.Succeeded) ? MutationOutcome.Failed : MutationOutcome.Applied
                    : MutationOutcome.AlreadySatisfied,
                results => $"Attempted {results.Count(result => result.Attempted)} approved field write(s); {results.Count(result => result.Attempted && !result.Succeeded)} failed.");
        }

        private static void WriteWebParts(
            Web targetWeb,
            PublishingPageMigrationPackage package,
            IDictionary<Guid, ListMaterializationReceipt> listReceipts,
            MigrationExecutionRecorder recorder)
        {
            if (package.Snapshot.WebParts.Count == 0)
            {
                recorder.RecordAlreadySatisfied("page.webparts", "The approved source snapshot has no shared Web Parts.");
                return;
            }

            var actions = package.Plan.WebPartActions.ToDictionary(value => value.SourceWebPartId);
            var bindings = package.Snapshot.ListWebPartBindings.ToDictionary(value => value.SourceWebPartId);
            foreach (var webPart in package.Snapshot.WebParts)
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
                    package.Plan.Replacements);
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
