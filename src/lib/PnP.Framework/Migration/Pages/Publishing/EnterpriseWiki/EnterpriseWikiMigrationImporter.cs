using Microsoft.SharePoint.Client;
using PnP.Framework.Entities;
using PnP.Framework.Migration.Pages.Content;
using PnP.Framework.Migration.Pages.Fields;
using PnP.Framework.Migration.Pages.Lifecycle;
using PnP.Framework.Migration.Pages.Publishing.Lifecycle;
using PnP.Framework.Migration.Pages.Publishing.Packaging;
using PnP.Framework.Migration.Pages.References;
using PnP.Framework.Migration.Pages.Publishing.Verification;
using System;
using System.Collections.Generic;
using System.Linq;

namespace PnP.Framework.Migration.Pages.Publishing.EnterpriseWiki
{
    public sealed class EnterpriseWikiMigrationImporter
    {
        public PublishingPageImportReceipt Import(
            ClientContext targetContext,
            PublishingPageMigrationPackage package,
            string approvedPlanDigest)
        {
            if (targetContext == null)
            {
                throw new ArgumentNullException(nameof(targetContext));
            }

            PublishingPagePackageValidator.ValidateMigration(package);
            EnterpriseWikiImportPlanValidator.Validate(package);
            if (package.State != PublishingPagePackageState.ApprovalReady || !package.Plan.IsExecutable)
            {
                throw new InvalidOperationException("The publishing-page package is not approval-ready.");
            }

            if (string.IsNullOrWhiteSpace(approvedPlanDigest)
                || !string.Equals(approvedPlanDigest, package.PlanDigest, StringComparison.OrdinalIgnoreCase))
            {
                throw new InvalidOperationException("The approved plan digest does not match the sealed publishing-page package.");
            }

            var startedAt = DateTimeOffset.UtcNow;
            var targetWeb = targetContext.Web;
            targetContext.Load(targetWeb, web => web.Url, web => web.ServerRelativeUrl);
            targetContext.ExecuteQueryRetry();
            if (!PagePath.UriEquals(targetWeb.Url, package.Plan.TargetWebUrl))
            {
                throw new InvalidOperationException($"The target connection points to '{targetWeb.Url}', but the approved plan targets '{package.Plan.TargetWebUrl}'.");
            }

            var preflightBlockers = new List<string>();
            var importWarnings = new List<string>();
            var freshProbe = EnterpriseWikiTargetInspector.Inspect(
                targetContext,
                package.Plan.TargetPageServerRelativeUrl,
                package.Plan.DependencyActions,
                package.Plan.TargetLifecycle,
                preflightBlockers);
            if (preflightBlockers.Count > 0)
            {
                throw new InvalidOperationException("Fresh target preflight failed: " + string.Join(" ", preflightBlockers));
            }

            if (!string.Equals(freshProbe.PageContentTypeId, package.Plan.TargetProbe.PageContentTypeId, StringComparison.OrdinalIgnoreCase)
                || !string.Equals(freshProbe.PageLayoutUrl, package.Plan.TargetProbe.PageLayoutUrl, StringComparison.OrdinalIgnoreCase))
            {
                throw new InvalidOperationException("The target Enterprise Wiki content type or layout changed after approval.");
            }

            var materialized = PageReferenceMaterializer.Materialize(
                targetContext,
                package.Snapshot.Dependencies,
                package.Plan.DependencyActions);
            var rewrittenContent = PageTextTransformer.Rewrite(
                package.Snapshot.PublishingPageContent,
                package.Plan.Replacements);
            var pages = targetWeb.GetPagesLibrary();
            if (pages == null)
            {
                throw new InvalidOperationException("The target publishing Pages library is unavailable.");
            }

            targetContext.Load(pages, list => list.EnableModeration, list => list.ForceCheckout);
            targetContext.Load(pages.RootFolder, folder => folder.ServerRelativeUrl);
            targetContext.ExecuteQueryRetry();
            var targetDirectory = PagePath.GetDirectoryName(package.Plan.TargetPageServerRelativeUrl);
            if (!string.Equals(targetDirectory, pages.RootFolder.ServerRelativeUrl, StringComparison.OrdinalIgnoreCase))
            {
                throw new NotSupportedException("The Enterprise Wiki import profile supports pages in the root of the target Pages library only.");
            }

            var targetFileName = PagePath.GetFileName(package.Plan.TargetPageServerRelativeUrl);
            targetWeb.AddPublishingPage(
                targetFileName,
                package.Plan.PageLayoutName,
                package.Snapshot.Source.Title,
                false);
            var targetFile = targetWeb.GetFileByServerRelativePath(ResourcePath.FromDecodedUrl(package.Plan.TargetPageServerRelativeUrl));
            var targetItem = targetFile.ListItemAllFields;
            targetContext.Load(targetFile, file => file.Exists, file => file.CheckOutType);
            targetContext.Load(targetItem, item => item.Id);
            targetContext.ExecuteQueryRetry();
            if (pages.ForceCheckout && targetFile.CheckOutType == CheckOutType.None)
            {
                targetFile.CheckOut();
                targetContext.ExecuteQueryRetry();
            }

            targetItem["Title"] = package.Snapshot.Source.Title;
            targetItem["PublishingPageContent"] = rewrittenContent;
            targetItem.Update();
            targetContext.ExecuteQueryRetry();
            var fieldResults = PageFieldWriter.Apply(
                targetContext,
                targetItem,
                package.Snapshot.Fields,
                package.Plan.FieldActions,
                package.Plan.Replacements,
                importWarnings);

            foreach (var webPart in package.Snapshot.WebParts)
            {
                targetWeb.AddWebPartToWebPartPage(package.Plan.TargetPageServerRelativeUrl, new WebPartEntity
                {
                    WebPartIndex = webPart.ZoneIndex,
                    WebPartTitle = webPart.Title,
                    WebPartZone = webPart.ZoneId,
                    WebPartXml = PageTextTransformer.Rewrite(webPart.ExportXml, package.Plan.Replacements)
                });
            }

            targetContext.Load(targetFile, file => file.CheckOutType);
            targetContext.ExecuteQueryRetry();
            var plannedFieldFailure = fieldResults.Any(result => result.Attempted && !result.Succeeded);
            if (targetFile.CheckOutType != CheckOutType.None)
            {
                var checkinType = package.Plan.TargetLifecycle == PublishingPageTargetLifecycle.Published && !plannedFieldFailure
                    ? CheckinType.MajorCheckIn
                    : CheckinType.MinorCheckIn;
                targetFile.CheckIn("PnP publishing-page import", checkinType);
                targetContext.ExecuteQueryRetry();
            }

            if (package.Plan.TargetLifecycle == PublishingPageTargetLifecycle.Published && !plannedFieldFailure)
            {
                targetFile.Publish("PnP publishing-page import");
                targetContext.ExecuteQueryRetry();
                if (pages.EnableModeration)
                {
                    targetFile.Approve("PnP publishing-page import");
                    targetContext.ExecuteQueryRetry();
                }
            }
            else if (plannedFieldFailure)
            {
                importWarnings.Add("One or more planned field updates failed. The page was not published.");
            }

            return PublishingPageImportVerifier.Verify(
                targetContext,
                package,
                approvedPlanDigest,
                startedAt,
                materialized,
                fieldResults,
                importWarnings,
                EnterpriseWikiMigrationProfile.IsContentType);
        }
    }
}
