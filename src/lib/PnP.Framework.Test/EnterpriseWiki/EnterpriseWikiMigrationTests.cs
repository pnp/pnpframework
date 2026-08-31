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
        public void ExportValidationDetectsSnapshotMutation()
        {
            var snapshot = CreateSnapshot();
            var export = new EnterpriseWikiExportPackage
            {
                ExportedAtUtc = DateTimeOffset.UtcNow,
                Snapshot = snapshot,
                SnapshotDigest = EnterpriseWikiPackageSerializer.ComputeSnapshotDigest(snapshot)
            };

            EnterpriseWikiPackageSerializer.ValidateExport(export);
            export.Snapshot.PublishingPageContent = "<p>changed</p>";

            Assert.ThrowsException<InvalidDataException>(() => EnterpriseWikiPackageSerializer.ValidateExport(export));
        }

        [TestMethod]
        public void MigrationValidationDetectsPlanMutation()
        {
            var package = CreateMigrationPackage();

            EnterpriseWikiPackageSerializer.ValidateMigration(package);
            package.Plan.TargetPageServerRelativeUrl = "/sites/target/Pages/changed.aspx";

            Assert.ThrowsException<InvalidDataException>(() => EnterpriseWikiPackageSerializer.ValidateMigration(package));
        }

        [TestMethod]
        public void LifecycleRuleMapsRealR11DraftAndE05PublishedEvidence()
        {
            var r11 = new EnterpriseWikiLifecycleSnapshot
            {
                CheckOutType = "Online",
                Level = "Draft",
                ModerationStatus = 3
            };
            var e05 = new EnterpriseWikiLifecycleSnapshot
            {
                CheckOutType = "None",
                Level = "Published",
                ModerationStatus = 0
            };

            Assert.AreEqual(EnterpriseWikiTargetLifecycle.Draft, EnterpriseWikiMigrationService.DeriveTargetLifecycle(r11));
            Assert.AreEqual(EnterpriseWikiTargetLifecycle.Published, EnterpriseWikiMigrationService.DeriveTargetLifecycle(e05));
            Assert.AreEqual(EnterpriseWikiTargetLifecycle.Draft, EnterpriseWikiMigrationService.DeriveTargetLifecycle(new EnterpriseWikiLifecycleSnapshot
            {
                CheckOutType = "Online",
                Level = "Published",
                ModerationStatus = 0
            }));
            Assert.AreEqual(EnterpriseWikiTargetLifecycle.Draft, EnterpriseWikiMigrationService.DeriveTargetLifecycle(null));
        }

        [TestMethod]
        public void ReportIncludesEveryCapturedFieldAndItsPlanDisposition()
        {
            var package = CreateMigrationPackage();

            var report = EnterpriseWikiPackageSerializer.BuildReport(package);

            StringAssert.Contains(report, "OOCLReference");
            StringAssert.Contains(report, "Custom recovery field");
            StringAssert.Contains(report, "EvidenceOnly");
            StringAssert.Contains(report, "rawValueJson");
            StringAssert.Contains(report, "Only Published maps to Published");
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

        private static EnterpriseWikiSnapshot CreateSnapshot()
        {
            return new EnterpriseWikiSnapshot
            {
                CapturePolicy = new EnterpriseWikiExportOptions
                {
                    SourcePageServerRelativeUrl = "/sites/source/Pages/source.aspx"
                },
                Source = new EnterpriseWikiPageIdentity
                {
                    WebUrl = "https://source.sharepoint.com/sites/source",
                    WebServerRelativeUrl = "/sites/source",
                    PageServerRelativeUrl = "/sites/source/Pages/source.aspx",
                    ContentTypeId = BuiltInContentTypeId.EnterpriseWikiPage,
                    ContentTypeName = "Enterprise Wiki Page",
                    Title = "Source",
                    PageLayoutUrl = "https://source.sharepoint.com/_catalogs/masterpage/EnterpriseWiki.aspx"
                },
                PublishingPageContent = "<p>source</p>",
                PublishingPageContentSha256 = EnterpriseWikiPackageSerializer.ComputeSha256("<p>source</p>"),
                Fields = new List<EnterpriseWikiFieldValueSnapshot>
                {
                    new EnterpriseWikiFieldValueSnapshot
                    {
                        Id = Guid.Parse("20d0d2ea-fd8e-4e91-a549-70a48e8932ef"),
                        InternalName = "OOCLReference",
                        Title = "Custom recovery field",
                        TypeAsString = "Text",
                        SchemaXml = "<Field Name=\"OOCLReference\" Type=\"Text\" />",
                        HasValue = true,
                        Kind = EnterpriseWikiFieldValueKind.Unsupported,
                        RawType = "Contoso.CustomFieldValue",
                        RawValue = "OOCL-42",
                        RawValueJson = "{\"reference\":\"OOCL-42\"}",
                        CaptureStatus = EnterpriseWikiCaptureStatus.CapturedWithLimitations
                    }
                },
                Security = new EnterpriseWikiSecuritySnapshot(),
                Lifecycle = new EnterpriseWikiLifecycleSnapshot
                {
                    CheckOutType = "Online",
                    Level = "Draft",
                    ModerationStatus = 3
                },
                SourceFence = new EnterpriseWikiSourceFence()
            };
        }

        private static EnterpriseWikiMigrationPackage CreateMigrationPackage()
        {
            var snapshot = CreateSnapshot();
            var snapshotDigest = EnterpriseWikiPackageSerializer.ComputeSnapshotDigest(snapshot);
            var plan = new EnterpriseWikiMigrationPlan
            {
                SourceSnapshotDigest = snapshotDigest,
                SourceWebUrl = snapshot.Source.WebUrl,
                SourcePageServerRelativeUrl = snapshot.Source.PageServerRelativeUrl,
                TargetWebUrl = "https://target.sharepoint.com/sites/target",
                TargetWebServerRelativeUrl = "/sites/target",
                TargetPageServerRelativeUrl = "/sites/target/Pages/source.aspx",
                PageLayoutName = "EnterpriseWiki",
                TargetLifecycle = EnterpriseWikiTargetLifecycle.Draft,
                LifecycleReason = "The source file level is 'Draft', so the target will remain Draft.",
                PlanningPolicy = new EnterpriseWikiPlanningOptions
                {
                    TargetPageServerRelativeUrl = "/sites/target/Pages/source.aspx"
                },
                TargetProbe = new EnterpriseWikiTargetProbe(),
                FieldActions = new List<EnterpriseWikiFieldAction>
                {
                    new EnterpriseWikiFieldAction
                    {
                        SourceInternalName = "OOCLReference",
                        TargetInternalName = "OOCLReference",
                        Disposition = EnterpriseWikiFieldDisposition.EvidenceOnly,
                        Reason = "The field is retained for a future mapper."
                    }
                },
                ExpectedPublishingPageContentSha256 = snapshot.PublishingPageContentSha256
            };
            return new EnterpriseWikiMigrationPackage
            {
                PlannedAtUtc = DateTimeOffset.UtcNow,
                ExportedAtUtc = DateTimeOffset.UtcNow.AddMinutes(-1),
                State = EnterpriseWikiPackageState.ApprovalReady,
                Snapshot = snapshot,
                Plan = plan,
                SnapshotDigest = snapshotDigest,
                PlanDigest = EnterpriseWikiPackageSerializer.ComputePlanDigest(plan),
                Report = new EnterpriseWikiCustomerReport
                {
                    Summary = "Test report"
                }
            };
        }
    }
}
