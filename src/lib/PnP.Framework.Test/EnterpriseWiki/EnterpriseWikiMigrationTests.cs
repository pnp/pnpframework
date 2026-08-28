using Microsoft.VisualStudio.TestTools.UnitTesting;
using PnP.Framework.EnterpriseWiki;
using System;
using System.Collections.Generic;
using System.IO;

namespace PnP.Framework.Test.EnterpriseWiki
{
    [TestClass]
    public class EnterpriseWikiMigrationTests
    {
        [TestMethod]
        public void ContentTypeClassificationExcludesProjectPages()
        {
            Assert.IsTrue(EnterpriseWikiMigrationService.IsEnterpriseWikiContentType(BuiltInContentTypeId.EnterpriseWikiPage + "001122"));
            Assert.IsFalse(EnterpriseWikiMigrationService.IsEnterpriseWikiContentType(BuiltInContentTypeId.ProjectPage + "001122"));
            Assert.IsFalse(EnterpriseWikiMigrationService.IsEnterpriseWikiContentType("0x010100C568DB52D9"));
        }

        [TestMethod]
        public void ContentRewriteUsesLongestCaseInsensitiveMappingFirst()
        {
            var replacements = new[]
            {
                new EnterpriseWikiTextReplacement
                {
                    Source = "https://source.sharepoint.com/sites/source",
                    Target = "https://target.sharepoint.com/sites/target"
                },
                new EnterpriseWikiTextReplacement
                {
                    Source = "https://source.sharepoint.com",
                    Target = "https://target.sharepoint.com"
                }
            };

            var actual = EnterpriseWikiMigrationService.RewriteContent(
                "<a href=\"HTTPS://SOURCE.SHAREPOINT.COM/sites/source/Pages/A.aspx\">A</a>",
                replacements);

            Assert.AreEqual("<a href=\"https://target.sharepoint.com/sites/target/Pages/A.aspx\">A</a>", actual);
        }

        [TestMethod]
        public void PackageValidationDetectsSnapshotMutation()
        {
            var snapshot = new EnterpriseWikiSnapshot
            {
                Source = new EnterpriseWikiPageIdentity
                {
                    WebUrl = "https://source.sharepoint.com/sites/source",
                    PageServerRelativeUrl = "/sites/source/Pages/source.aspx"
                },
                PublishingPageContent = "<p>source</p>",
                PublishingPageContentSha256 = EnterpriseWikiPackageSerializer.ComputeSha256("<p>source</p>")
            };
            var snapshotDigest = EnterpriseWikiPackageSerializer.ComputeSnapshotDigest(snapshot);
            var plan = new EnterpriseWikiMigrationPlan
            {
                SourceSnapshotDigest = snapshotDigest,
                TargetWebUrl = "https://target.sharepoint.com/sites/target",
                TargetPageServerRelativeUrl = "/sites/target/Pages/source.aspx",
                TargetProbe = new EnterpriseWikiTargetProbe(),
                Blockers = new List<string>(),
                Warnings = new List<string>()
            };
            var package = new EnterpriseWikiMigrationPackage
            {
                CreatedAtUtc = DateTimeOffset.UtcNow,
                State = EnterpriseWikiPackageState.ApprovalReady,
                Snapshot = snapshot,
                Plan = plan,
                SnapshotDigest = snapshotDigest,
                PlanDigest = EnterpriseWikiPackageSerializer.ComputePlanDigest(plan)
            };

            EnterpriseWikiPackageSerializer.Validate(package);
            package.Snapshot.PublishingPageContent = "<p>changed</p>";

            Assert.ThrowsException<InvalidDataException>(() => EnterpriseWikiPackageSerializer.Validate(package));
        }

        [TestMethod]
        public void WebPartPortabilityBlocksSourceListBindingsAndKnownUnsupportedTypes()
        {
            const string listView = @"<webParts><webPart xmlns=""http://schemas.microsoft.com/WebPart/v3""><metaData><type name=""Microsoft.SharePoint.WebPartPages.XsltListViewWebPart"" /></metaData><data><properties><property name=""ListId"">58a84d5d-b1ee-4da0-a49b-7e597ee8ae35</property></properties></data></webPart></webParts>";
            const string rss = @"<webParts><webPart xmlns=""http://schemas.microsoft.com/WebPart/v3""><metaData><type name=""Microsoft.SharePoint.Portal.WebControls.RSSAggregatorWebPart"" /></metaData></webPart></webParts>";
            const string scriptEditor = @"<webParts><webPart xmlns=""http://schemas.microsoft.com/WebPart/v3""><metaData><type name=""Microsoft.SharePoint.WebPartPages.ScriptEditorWebPart"" /></metaData></webPart></webParts>";

            StringAssert.Contains(EnterpriseWikiMigrationService.GetWebPartMigrationBlocker(listView), "reviewed target-list");
            StringAssert.Contains(EnterpriseWikiMigrationService.GetWebPartMigrationBlocker(rss), "not supported");
            Assert.IsNull(EnterpriseWikiMigrationService.GetWebPartMigrationBlocker(scriptEditor));
        }
    }
}
