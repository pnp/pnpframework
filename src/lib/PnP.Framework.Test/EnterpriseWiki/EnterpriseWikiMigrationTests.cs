using Microsoft.VisualStudio.TestTools.UnitTesting;
using PnP.Framework.Migration.Pages;
using PnP.Framework.Migration.Pages.Capture;
using PnP.Framework.Migration.Pages.Content;
using PnP.Framework.Migration.Pages.Publishing.EnterpriseWiki;
using PnP.Framework.Migration.Pages.Fields;
using PnP.Framework.Migration.Pages.Lifecycle;
using PnP.Framework.Migration.Pages.Publishing.Capture;
using PnP.Framework.Migration.Pages.Publishing.Lifecycle;
using PnP.Framework.Migration.Pages.Publishing.Packaging;
using PnP.Framework.Migration.Pages.Publishing.Planning;
using PnP.Framework.Migration.Pages.Planning;
using PnP.Framework.Migration.Pages.Publishing.Reporting;
using PnP.Framework.Migration.Pages.Security;
using PnP.Framework.Migration.Pages.Publishing.Verification;
using PnP.Framework.Migration.Pages.ClassicWebParts;
using PnP.Framework.Migration.Pages.Publishing.Layouts;
using PnP.Framework.Migration.Execution;
using PnP.Framework.Migration.Evidence;
using PnP.Framework.Migration.Packaging;
using PnP.Framework.Migration.Schema.ContentTypes;
using PnP.Framework.Migration.Schema.Fields;
using PnP.Framework.Migration.Verification;
using Microsoft.SharePoint.Client;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace PnP.Framework.Test.EnterpriseWiki
{
    [TestClass]
    public class EnterpriseWikiMigrationTests
    {
        [TestMethod]
        public void ContentTypeClassificationExcludesProjectPages()
        {
            Assert.IsTrue(EnterpriseWikiPageDiscovery.IsEnterpriseWikiContentType(BuiltInContentTypeId.EnterpriseWikiPage + "001122"));
            Assert.IsFalse(EnterpriseWikiPageDiscovery.IsEnterpriseWikiContentType(BuiltInContentTypeId.ProjectPage + "001122"));
            Assert.IsFalse(EnterpriseWikiPageDiscovery.IsEnterpriseWikiContentType("0x010100C568DB52D9"));
        }

        [TestMethod]
        public void ContentRewriteUsesLongestCaseInsensitiveMappingFirst()
        {
            var replacements = new[]
            {
                new PageTextReplacement
                {
                    Source = "https://source.sharepoint.com/sites/source",
                    Target = "https://target.sharepoint.com/sites/target"
                },
                new PageTextReplacement
                {
                    Source = "https://source.sharepoint.com",
                    Target = "https://target.sharepoint.com"
                }
            };

            var actual = PageTextTransformer.Rewrite(
                "<a href=\"HTTPS://SOURCE.SHAREPOINT.COM/sites/source/Pages/A.aspx\">A</a>",
                replacements);

            Assert.AreEqual("<a href=\"https://target.sharepoint.com/sites/target/Pages/A.aspx\">A</a>", actual);
        }

        [TestMethod]
        public void ExportValidationDetectsSnapshotMutation()
        {
            var snapshot = CreateSnapshot();
            var export = new PublishingPageExportPackage
            {
                ExportedAtUtc = DateTimeOffset.UtcNow,
                Snapshot = snapshot,
                SnapshotDigest = PublishingPageDigest.ComputeSnapshotDigest(snapshot)
            };

            PublishingPagePackageValidator.ValidateExport(export);
            export.Snapshot.PublishingPageContent = "<p>changed</p>";

            Assert.ThrowsException<InvalidDataException>(() => PublishingPagePackageValidator.ValidateExport(export));
        }

        [TestMethod]
        public void ExportValidationDetectsPageLayoutByteMutation()
        {
            var snapshot = CreateSnapshot();
            var export = new PublishingPageExportPackage
            {
                ExportedAtUtc = DateTimeOffset.UtcNow,
                Snapshot = snapshot,
                SnapshotDigest = PublishingPageDigest.ComputeSnapshotDigest(snapshot)
            };

            export.Snapshot.Layout.ContentBase64 = Convert.ToBase64String(Encoding.UTF8.GetBytes("changed"));

            Assert.ThrowsException<InvalidDataException>(() => PublishingPagePackageValidator.ValidateExport(export));
        }

        [TestMethod]
        public void MigrationValidationDetectsPlanMutation()
        {
            var package = CreateMigrationPackage();

            PublishingPagePackageValidator.ValidateMigration(package);
            package.Plan.TargetPageServerRelativeUrl = "/sites/target/Pages/changed.aspx";

            Assert.ThrowsException<InvalidDataException>(() => PublishingPagePackageValidator.ValidateMigration(package));
        }

        [TestMethod]
        public void LifecycleRuleMapsRealR11DraftAndE05PublishedEvidence()
        {
            var r11 = new PageLifecycleSnapshot
            {
                CheckOutType = "Online",
                Level = "Draft",
                ModerationStatus = 3
            };
            var e05 = new PageLifecycleSnapshot
            {
                CheckOutType = "None",
                Level = "Published",
                ModerationStatus = 0
            };

            Assert.AreEqual(PublishingPageTargetLifecycle.Draft, PublishingPageLifecyclePolicy.DeriveTargetLifecycle(r11));
            Assert.AreEqual(PublishingPageTargetLifecycle.Published, PublishingPageLifecyclePolicy.DeriveTargetLifecycle(e05));
            Assert.AreEqual(PublishingPageTargetLifecycle.Draft, PublishingPageLifecyclePolicy.DeriveTargetLifecycle(new PageLifecycleSnapshot
            {
                CheckOutType = "Online",
                Level = "Published",
                ModerationStatus = 0
            }));
            Assert.AreEqual(PublishingPageTargetLifecycle.Draft, PublishingPageLifecyclePolicy.DeriveTargetLifecycle(null));
        }

        [TestMethod]
        public void ReportIncludesEveryCapturedFieldAndItsPlanDisposition()
        {
            var package = CreateMigrationPackage();

            var report = PublishingPageMigrationReportBuilder.Build(package);

            StringAssert.Contains(report, "OOCLReference");
            StringAssert.Contains(report, "Custom recovery field");
            StringAssert.Contains(report, "EvidenceOnly");
            StringAssert.Contains(report, "rawValueJson");
            StringAssert.Contains(report, "Only an unconflicted Published");
            StringAssert.Contains(report, "snapshot.sourceProfile");
            StringAssert.Contains(report, "snapshot.layout.customizedPageStatus");
            StringAssert.Contains(report, "Page Layout materialization plan");
            StringAssert.Contains(report, "Page Layout target admission");
        }

        [TestMethod]
        public void WebPartPortabilityBlocksSourceListBindingsAndKnownUnsupportedTypes()
        {
            const string listView = @"<webParts><webPart xmlns=""http://schemas.microsoft.com/WebPart/v3""><metaData><type name=""Microsoft.SharePoint.WebPartPages.XsltListViewWebPart"" /></metaData><data><properties><property name=""ListId"">58a84d5d-b1ee-4da0-a49b-7e597ee8ae35</property></properties></data></webPart></webParts>";
            const string rss = @"<webParts><webPart xmlns=""http://schemas.microsoft.com/WebPart/v3""><metaData><type name=""Microsoft.SharePoint.Portal.WebControls.RSSAggregatorWebPart"" /></metaData></webPart></webParts>";
            const string scriptEditor = @"<webParts><webPart xmlns=""http://schemas.microsoft.com/WebPart/v3""><metaData><type name=""Microsoft.SharePoint.WebPartPages.ScriptEditorWebPart"" /></metaData></webPart></webParts>";

            StringAssert.Contains(EnterpriseWikiWebPartPolicy.GetBlocker(listView), "reviewed target-list");
            StringAssert.Contains(EnterpriseWikiWebPartPolicy.GetBlocker(rss), "not supported");
            Assert.IsNull(EnterpriseWikiWebPartPolicy.GetBlocker(scriptEditor));
        }

        [TestMethod]
        public void SerializerRoundTripsTheGenericPublishingPageContract()
        {
            var package = CreateMigrationPackage();

            var json = PublishingPagePackageSerializer.Serialize(package);
            var roundTripped = PublishingPagePackageSerializer.Deserialize<PublishingPageMigrationPackage>(json);

            PublishingPagePackageValidator.ValidateMigration(roundTripped);
            Assert.AreEqual("EnterpriseWiki", roundTripped.Snapshot.SourceProfile);
            Assert.AreEqual("https://source.sharepoint.com/_catalogs/masterpage/EnterpriseWiki.aspx", roundTripped.Snapshot.Layout.Url);
            Assert.AreEqual(package.Snapshot.Layout.Bytes.Sha256, roundTripped.Snapshot.Layout.Bytes.Sha256);
            Assert.AreEqual("PublishingPageContent", roundTripped.Snapshot.Layout.Controls.Single().FieldName);
            Assert.AreEqual(package.PlanDigest, roundTripped.PlanDigest);
            Assert.AreEqual(1, roundTripped.Plan.RuntimeVerification.Requirements.Count);
            Assert.AreEqual(RuntimeVerificationRequirementKind.AuthoredDomEquality, roundTripped.Plan.RuntimeVerification.Requirements[0].Kind);
        }

        [TestMethod]
        public void RuntimeVerificationRequirementsAreSealedByThePlanDigest()
        {
            var package = CreateMigrationPackage();

            package.Plan.RuntimeVerification.Requirements[0].Description = "changed after approval";

            Assert.ThrowsException<InvalidDataException>(() => PublishingPagePackageValidator.ValidateMigration(package));
        }

        [TestMethod]
        public void ImportReturnsAZeroMutationReceiptWhenThePlanDigestWasNotApproved()
        {
            var package = CreateMigrationPackage();
            var journal = new InMemoryMigrationExecutionJournal();
            using (var context = new ClientContext(package.Plan.TargetWebUrl))
            {
                var receipt = new EnterpriseWikiMigrationImporter().Import(context, package, "not-approved", journal);

                Assert.AreEqual(MigrationExecutionStatus.NotStarted, receipt.ExecutionStatus);
                Assert.AreEqual("PlanDigestNotApproved", receipt.AdmissionFailure.Code);
                Assert.IsFalse(receipt.MutationStarted);
                Assert.AreEqual(StorageVerificationStatus.NotRun, receipt.StorageVerificationStatus);
                Assert.AreEqual(RuntimeVerificationStatus.NotRun, receipt.RuntimeVerificationStatus);
                Assert.AreEqual(MigrationAcceptanceStatus.Rejected, receipt.AcceptanceStatus);
                Assert.AreEqual(0, journal.Intents.Count);
                Assert.AreEqual(1, journal.ExecutionStates.Count);
            }
        }

        [TestMethod]
        public void LayoutMarkupParserCapturesFieldsZonesRegistrationsAndEncodedResources()
        {
            const string markup = @"<%@ Register TagPrefix=""PublishingWebControls"" Namespace=""Microsoft.SharePoint.Publishing.WebControls"" Assembly=""Microsoft.SharePoint.Publishing"" %>
<PublishingWebControls:RichHtmlField ID=""PageContent"" FieldName=""PublishingPageContent"" runat=""server"" />
<WebPartPages:WebPartZone ID=""Main"" runat=""server"" />
<SharePoint:CssRegistration Name=""<% $SPUrl:~sitecollection/Style Library/Contoso/site.css %>"" runat=""server"" />
&lt;script src=&quot;~site/SiteAssets/Contoso/app.js&quot;&gt;&lt;/script&gt;";

            var parsed = PublishingPageLayoutMarkupParser.Parse(markup);

            Assert.AreEqual(1, parsed.Registrations.Count);
            Assert.IsTrue(parsed.RequiredFieldNames.Contains("Title", StringComparer.OrdinalIgnoreCase));
            Assert.IsTrue(parsed.RequiredFieldNames.Contains("PublishingPageContent", StringComparer.OrdinalIgnoreCase));
            Assert.AreEqual("Main", parsed.Zones.Single().Id);
            Assert.IsTrue(parsed.ResourceReferences.Any(value => value.Value == "~sitecollection/Style Library/Contoso/site.css"));
            Assert.IsTrue(parsed.ResourceReferences.Any(value => value.Value == "~site/SiteAssets/Contoso/app.js"));
        }

        [TestMethod]
        public void FieldSchemaCanonicalizerIgnoresStorageSlotsAndRebindsTaxonomy()
        {
            const string left = @"<Field ID=""{11111111-1111-1111-1111-111111111111}"" Name=""Category"" Type=""TaxonomyFieldType"" SourceID=""source-a"" ColName=""nvarchar1"" RowOrdinal=""1""><Customization><ArrayOfProperty><Property><Name>SspId</Name><Value>{22222222-2222-2222-2222-222222222222}</Value></Property><Property><Name>TermSetId</Name><Value>33333333-3333-3333-3333-333333333333</Value></Property><Property><Name>TextField</Name><Value>44444444-4444-4444-4444-444444444444</Value></Property></ArrayOfProperty></Customization></Field>";
            const string right = @"<Field RowOrdinal=""99"" ColName=""nvarchar42"" SourceID=""source-b"" Type=""TaxonomyFieldType"" Name=""Category"" ID=""{11111111-1111-1111-1111-111111111111}""><Customization><ArrayOfProperty><Property><Name>SspId</Name><Value>22222222-2222-2222-2222-222222222222</Value></Property><Property><Name>TermSetId</Name><Value>{33333333-3333-3333-3333-333333333333}</Value></Property><Property><Name>TextField</Name><Value>{44444444-4444-4444-4444-444444444444}</Value></Property></ArrayOfProperty></Customization></Field>";

            Assert.AreEqual(FieldSchemaCanonicalizer.PortableDigest(left), FieldSchemaCanonicalizer.PortableDigest(right));

            var rewritten = FieldSchemaCanonicalizer.RewriteForTarget(
                left,
                Guid.Parse("aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa"),
                Guid.Parse("bbbbbbbb-bbbb-bbbb-bbbb-bbbbbbbbbbbb"),
                Guid.Parse("cccccccc-cccc-cccc-cccc-cccccccccccc"));

            Assert.IsFalse(rewritten.Contains("ColName"));
            Assert.IsFalse(rewritten.Contains("RowOrdinal"));
            StringAssert.Contains(rewritten, "aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa");
            StringAssert.Contains(rewritten, "bbbbbbbb-bbbb-bbbb-bbbb-bbbbbbbbbbbb");
            StringAssert.Contains(rewritten, "cccccccc-cccc-cccc-cccc-cccccccccccc");
        }

        [TestMethod]
        public void ContentTypeSchemaPlannerCreatesScalarClosureAndBlocksUnmappedTaxonomy()
        {
            var scalarId = Guid.Parse("11111111-1111-1111-1111-111111111111");
            var taxonomyId = Guid.Parse("22222222-2222-2222-2222-222222222222");
            var schema = new ContentTypeSchemaSnapshot
            {
                EvidenceState = ContentTypeSchemaEvidenceState.Readable,
                Availability = EvidenceAvailability.Captured,
                SourceWebUrl = "https://source.sharepoint.com/sites/source",
                ContentTypeId = "0x010100AA",
                Name = "Custom Page",
                ParentContentTypeId = "0x01010007FF3E057FA8AB4AA42FCB67B453FFC",
                ParentContentTypeName = "Page",
                RequiredFieldLinks = new List<ContentTypeFieldLinkSnapshot>
                {
                    new ContentTypeFieldLinkSnapshot { FieldId = scalarId, Name = "Activity", Role = FieldSchemaRole.DirectBinding },
                    new ContentTypeFieldLinkSnapshot { FieldId = taxonomyId, Name = "Category", Role = FieldSchemaRole.DirectBinding }
                },
                RequiredFieldClosure = new List<FieldSchemaSnapshot>
                {
                    Field(scalarId, "Activity", "Text", "<Field ID=\"{11111111-1111-1111-1111-111111111111}\" Name=\"Activity\" Type=\"Text\" ColName=\"nvarchar1\" />"),
                    new FieldSchemaSnapshot
                    {
                        Id = taxonomyId,
                        InternalName = "Category",
                        Title = "Category",
                        TypeAsString = "TaxonomyFieldType",
                        SchemaXml = "<Field ID=\"{22222222-2222-2222-2222-222222222222}\" Name=\"Category\" Type=\"TaxonomyFieldType\" />",
                        PortableSchemaSha256 = FieldSchemaCanonicalizer.PortableDigest("<Field ID=\"{22222222-2222-2222-2222-222222222222}\" Name=\"Category\" Type=\"TaxonomyFieldType\" />"),
                        Role = FieldSchemaRole.DirectBinding,
                        Taxonomy = new TaxonomyFieldBindingSnapshot
                        {
                            SourceTermStoreId = Guid.Parse("33333333-3333-3333-3333-333333333333"),
                            SourceTermSetId = Guid.Parse("44444444-4444-4444-4444-444444444444"),
                            HiddenTextFieldId = Guid.Parse("55555555-5555-5555-5555-555555555555")
                        }
                    }
                }
            };

            var plan = ContentTypeSchemaPlanner.CreateRequiredClosure(schema);

            Assert.AreEqual(ContentTypeMaterializationDisposition.Block, plan.Disposition);
            Assert.AreEqual(FieldSchemaMaterializationDisposition.CreateOrReuseOwned,
                plan.Fields.Single(value => value.FieldId == scalarId).Disposition);
            Assert.AreEqual(FieldSchemaMaterializationDisposition.Block,
                plan.Fields.Single(value => value.FieldId == taxonomyId).Disposition);
            Assert.IsFalse(plan.Fields.Single(value => value.FieldId == scalarId).TargetSchemaXml.Contains("ColName"));
        }

        [TestMethod]
        public void CustomLayoutPlanMapsExactSchemaResourcesAndDeterministicTargetBytes()
        {
            var layout = CreateCustomLayout();

            var plan = PublishingPageLayoutPlanFactory.Create(
                layout,
                new Uri("https://source.sharepoint.com/sites/source"),
                new Uri("https://target.sharepoint.com/sites/target"),
                new Uri("https://target.sharepoint.com/sites/target"),
                "EnterpriseWiki.aspx");

            Assert.AreEqual(PublishingPageLayoutMaterializationDisposition.CreateOwned, plan.Disposition);
            StringAssert.StartsWith(plan.TargetFileName, "pnp-custom-");
            StringAssert.EndsWith(plan.TargetFileName, ".aspx");
            Assert.AreEqual("/sites/target/_catalogs/masterpage/" + plan.TargetFileName, plan.TargetServerRelativeUrl);
            Assert.AreEqual(ContentTypeMaterializationDisposition.CreateOwned, plan.ContentTypeSchema.Disposition);
            Assert.AreEqual(FieldSchemaMaterializationDisposition.CreateOrReuseOwned, plan.ContentTypeSchema.Fields.Single().Disposition);
            Assert.AreEqual("/sites/target/SiteAssets/Contoso/app.js", plan.ResourceMaterializations.Single().TargetServerRelativeUrl);
            Assert.AreEqual("https://target.sharepoint.com/sites/target/SiteAssets/Contoso/app.js", plan.ResourceRewrites.Single().TargetReference);
            Assert.AreNotEqual(plan.SourceBytes.Sha256, plan.TargetBytes.Sha256);
        }

        [TestMethod]
        public void CustomLayoutAdmissionCreatesOnlyWhenSchemaResourcesAndTargetAreEligible()
        {
            var layout = CreateCustomLayout();
            var plan = PublishingPageLayoutPlanFactory.Create(
                layout,
                new Uri("https://source.sharepoint.com/sites/source"),
                new Uri("https://target.sharepoint.com/sites/target"),
                new Uri("https://target.sharepoint.com/sites/target"),
                "EnterpriseWiki.aspx");
            var probe = CreateEligibleCustomLayoutProbe(plan);

            var admission = PublishingPageLayoutTargetAdmissionEvaluator.Evaluate(plan, probe);

            Assert.IsTrue(admission.IsEligible);
            Assert.AreEqual(PublishingPageLayoutMaterializationDisposition.CreateOwned, admission.Disposition);
            Assert.AreEqual(ContentTypeMaterializationDisposition.CreateOwned, admission.ContentTypeSchema.Disposition);

            probe.Resources.Single().FileExists = true;
            probe.Resources.Single().ExistingBytesSha256 = new string('0', 64);
            var collision = PublishingPageLayoutTargetAdmissionEvaluator.Evaluate(plan, probe);
            Assert.IsFalse(collision.IsEligible);
            Assert.IsTrue(collision.Issues.Any(value => value.Code == "TargetLayoutResourceCollision"));
        }

        [TestMethod]
        public void ExportValidationRequiresOneEvidenceRecordPerLayoutResourceReference()
        {
            var snapshot = CreateSnapshot();
            snapshot.Layout = CreateCustomLayout();
            snapshot.Layout.ResourceArtifacts.Clear();
            var export = new PublishingPageExportPackage
            {
                ExportedAtUtc = DateTimeOffset.UtcNow,
                Snapshot = snapshot,
                SnapshotDigest = PublishingPageDigest.ComputeSnapshotDigest(snapshot)
            };

            Assert.ThrowsException<InvalidDataException>(() => PublishingPagePackageValidator.ValidateExport(export));
        }

        [TestMethod]
        public void DirectoryArtifactStoreRoundTripsAndDeduplicatesByDigest()
        {
            var root = Path.Combine(Path.GetTempPath(), "pnp-migration-artifacts-" + Guid.NewGuid().ToString("N"));
            try
            {
                var store = new DirectoryMigrationArtifactStore(root);
                var bytes = Encoding.UTF8.GetBytes("exact migration payload");
                ArtifactReference first;
                ArtifactReference second;
                using (var content = new MemoryStream(bytes, false))
                {
                    first = store.Put(content, "text/plain", "payload.txt");
                }

                using (var content = new MemoryStream(bytes, false))
                {
                    second = store.Put(content, "text/plain", "another-name.txt");
                }

                Assert.AreEqual(first.Sha256, second.Sha256);
                Assert.AreEqual(bytes.LongLength, first.Length);
                Assert.IsTrue(store.Contains(first.Sha256));
                using (var content = store.OpenRead(first.Sha256))
                using (var buffer = new MemoryStream())
                {
                    content.CopyTo(buffer);
                    CollectionAssert.AreEqual(bytes, buffer.ToArray());
                }
            }
            finally
            {
                if (Directory.Exists(root))
                {
                    Directory.Delete(root, true);
                }
            }
        }

        private static PublishingPageCaptureBundle CreateSnapshot()
        {
            return new PublishingPageCaptureBundle
            {
                SourceProfile = "EnterpriseWiki",
                CapturePolicy = new PageCaptureOptions
                {
                    SourcePageServerRelativeUrl = "/sites/source/Pages/source.aspx"
                },
                Source = new PageIdentity
                {
                    WebUrl = "https://source.sharepoint.com/sites/source",
                    WebServerRelativeUrl = "/sites/source",
                    PageServerRelativeUrl = "/sites/source/Pages/source.aspx",
                    ContentTypeId = BuiltInContentTypeId.EnterpriseWikiPage,
                    ContentTypeName = "Enterprise Wiki Page",
                    Title = "Source"
                },
                Layout = CreateStockLayout(),
                PublishingPageContent = "<p>source</p>",
                PublishingPageContentSha256 = PublishingPageDigest.ComputeSha256("<p>source</p>"),
                Fields = new List<PageFieldValueSnapshot>
                {
                    new PageFieldValueSnapshot
                    {
                        Id = Guid.Parse("20d0d2ea-fd8e-4e91-a549-70a48e8932ef"),
                        InternalName = "OOCLReference",
                        Title = "Custom recovery field",
                        TypeAsString = "Text",
                        SchemaXml = "<Field Name=\"OOCLReference\" Type=\"Text\" />",
                        HasValue = true,
                        Kind = PageFieldValueKind.Unsupported,
                        RawType = "Contoso.CustomFieldValue",
                        RawValue = "OOCL-42",
                        RawValueJson = "{\"reference\":\"OOCL-42\"}",
                        CaptureStatus = PageCaptureStatus.CapturedWithLimitations
                    }
                },
                Security = new PageSecuritySnapshot(),
                Lifecycle = new PageLifecycleSnapshot
                {
                    CheckOutType = "Online",
                    Level = "Draft",
                    ModerationStatus = 3
                },
                SourceFence = new SourcePageFence()
            };
        }

        private static PublishingPageLayoutSnapshot CreateStockLayout()
        {
            var bytes = Encoding.UTF8.GetBytes("<%@ Page %><PublishingWebControls:RichHtmlField FieldName=\"PublishingPageContent\" runat=\"server\" />");
            return new PublishingPageLayoutSnapshot
            {
                EvidenceState = PublishingPageLayoutEvidenceState.Readable,
                Availability = EvidenceAvailability.Captured,
                Url = "https://source.sharepoint.com/_catalogs/masterpage/EnterpriseWiki.aspx",
                ServerRelativeUrl = "/_catalogs/masterpage/EnterpriseWiki.aspx",
                FileName = "EnterpriseWiki.aspx",
                CustomizedPageStatus = 1,
                AssociatedContentTypeName = "Enterprise Wiki Page",
                AssociatedContentTypeId = BuiltInContentTypeId.EnterpriseWikiPage,
                Bytes = MigrationArtifact.Describe(bytes, "application/vnd.ms-aspx", "EnterpriseWiki.aspx"),
                ContentBase64 = Convert.ToBase64String(bytes),
                Controls = new List<PublishingPageLayoutControl>
                {
                    new PublishingPageLayoutControl
                    {
                        TagPrefix = "PublishingWebControls",
                        ControlName = "RichHtmlField",
                        FieldName = "PublishingPageContent"
                    }
                }
            };
        }

        private static PublishingPageLayoutSnapshot CreateCustomLayout()
        {
            const string authoredReference = "~site/SiteAssets/Contoso/app.js";
            var layoutBytes = Encoding.UTF8.GetBytes(
                "<%@ Page %><PublishingWebControls:TextField FieldName=\"Activity\" runat=\"server\" /><script src=\""
                + authoredReference
                + "\"></script>");
            var resourceBytes = Encoding.UTF8.GetBytes("console.log('source');");
            var fieldId = Guid.Parse("11111111-1111-1111-1111-111111111111");
            var field = Field(
                fieldId,
                "Activity",
                "Text",
                "<Field ID=\"{11111111-1111-1111-1111-111111111111}\" Name=\"Activity\" DisplayName=\"Activity\" Type=\"Text\" />");
            var reference = new PublishingPageLayoutResourceReference
            {
                Attribute = "src",
                Value = authoredReference
            };
            return new PublishingPageLayoutSnapshot
            {
                EvidenceState = PublishingPageLayoutEvidenceState.Readable,
                Availability = EvidenceAvailability.Captured,
                Url = "https://source.sharepoint.com/sites/source/_catalogs/masterpage/Custom.aspx",
                ServerRelativeUrl = "/sites/source/_catalogs/masterpage/Custom.aspx",
                FileName = "Custom.aspx",
                CustomizedPageStatus = 0,
                AssociatedContentTypeName = "Custom Publishing Page",
                AssociatedContentTypeId = "0x01010007FF3E057FA8AB4AA42FCB67B453FFC001122",
                Bytes = MigrationArtifact.Describe(layoutBytes, "application/vnd.ms-aspx", "Custom.aspx"),
                ContentBase64 = Convert.ToBase64String(layoutBytes),
                Controls = new List<PublishingPageLayoutControl>
                {
                    new PublishingPageLayoutControl
                    {
                        TagPrefix = "PublishingWebControls",
                        ControlName = "TextField",
                        FieldName = "Activity"
                    }
                },
                ResourceReferences = new List<PublishingPageLayoutResourceReference> { reference },
                ResourceArtifacts = new List<PublishingPageLayoutResourceSnapshot>
                {
                    new PublishingPageLayoutResourceSnapshot
                    {
                        Reference = reference,
                        EvidenceState = PublishingPageLayoutResourceEvidenceState.Readable,
                        ResolvedSourceUrl = "https://source.sharepoint.com/sites/source/SiteAssets/Contoso/app.js",
                        Artifact = MigrationArtifact.Describe(resourceBytes, "text/javascript", "app.js"),
                        ContentBase64 = Convert.ToBase64String(resourceBytes)
                    }
                },
                AssociatedContentTypeSchema = new ContentTypeSchemaSnapshot
                {
                    EvidenceState = ContentTypeSchemaEvidenceState.Readable,
                    Availability = EvidenceAvailability.Captured,
                    SourceWebUrl = "https://source.sharepoint.com/sites/source",
                    ContentTypeId = "0x01010007FF3E057FA8AB4AA42FCB67B453FFC001122",
                    Name = "Custom Publishing Page",
                    Description = "Custom page schema",
                    Group = "Custom Content Types",
                    ParentContentTypeId = "0x01010007FF3E057FA8AB4AA42FCB67B453FFC",
                    ParentContentTypeName = "Page",
                    RequiredFieldLinks = new List<ContentTypeFieldLinkSnapshot>
                    {
                        new ContentTypeFieldLinkSnapshot
                        {
                            FieldId = fieldId,
                            Name = "Activity",
                            Role = FieldSchemaRole.DirectBinding
                        }
                    },
                    RequiredFieldClosure = new List<FieldSchemaSnapshot> { field }
                }
            };
        }

        private static PublishingPageLayoutTargetProbe CreateEligibleCustomLayoutProbe(
            PublishingPageLayoutMaterializationPlan plan)
        {
            return new PublishingPageLayoutTargetProbe
            {
                TargetServerRelativeUrl = plan.TargetServerRelativeUrl,
                FileExists = false,
                CanAddAndCustomizePages = true,
                Availability = EvidenceAvailability.Captured,
                ContentTypeSchema = new ContentTypeTargetProbe
                {
                    ContentTypeId = plan.ContentTypeSchema.ContentTypeId,
                    ParentContentTypeAvailable = true,
                    ResolvedParentContentTypeId = plan.ContentTypeSchema.ParentContentTypeId,
                    CanManageContentTypes = true,
                    Availability = EvidenceAvailability.Captured
                },
                Resources = plan.ResourceMaterializations
                    .Where(value => value.Disposition == PublishingPageLayoutResourceMaterializationDisposition.CreateOrReuseOwned)
                    .Select(value => new PublishingPageLayoutResourceTargetProbe
                    {
                        TargetServerRelativeUrl = value.TargetServerRelativeUrl,
                        FileExists = false,
                        CanWrite = true,
                        Availability = EvidenceAvailability.Captured
                    })
                    .ToList()
            };
        }

        private static FieldSchemaSnapshot Field(Guid id, string name, string type, string schemaXml)
        {
            return new FieldSchemaSnapshot
            {
                Id = id,
                InternalName = name,
                Title = name,
                TypeAsString = type,
                SchemaXml = schemaXml,
                SchemaXmlSha256 = MigrationDigest.ComputeSha256(schemaXml),
                PortableSchemaSha256 = FieldSchemaCanonicalizer.PortableDigest(schemaXml),
                Role = FieldSchemaRole.DirectBinding
            };
        }

        private static PublishingPageMigrationPackage CreateMigrationPackage()
        {
            var snapshot = CreateSnapshot();
            var snapshotDigest = PublishingPageDigest.ComputeSnapshotDigest(snapshot);
            var layoutPlan = PublishingPageLayoutPlanFactory.Create(
                snapshot.Layout,
                new Uri(snapshot.Source.WebUrl),
                new Uri("https://target.sharepoint.com/sites/target"),
                new Uri("https://target.sharepoint.com/sites/target"),
                "EnterpriseWiki.aspx");
            var layoutProbe = new PublishingPageLayoutTargetProbe
            {
                TargetServerRelativeUrl = layoutPlan.TargetServerRelativeUrl,
                FileExists = true,
                Availability = EvidenceAvailability.Captured
            };
            var layoutAdmission = PublishingPageLayoutTargetAdmissionEvaluator.Evaluate(layoutPlan, layoutProbe);
            var plan = new PublishingPageMigrationPlan
            {
                SourceSnapshotDigest = snapshotDigest,
                SourceWebUrl = snapshot.Source.WebUrl,
                SourcePageServerRelativeUrl = snapshot.Source.PageServerRelativeUrl,
                TargetWebUrl = "https://target.sharepoint.com/sites/target",
                TargetWebServerRelativeUrl = "/sites/target",
                TargetPageServerRelativeUrl = "/sites/target/Pages/source.aspx",
                PageLayoutName = "EnterpriseWiki",
                TargetLifecycle = PublishingPageTargetLifecycle.Draft,
                LifecycleReason = "The source file level is 'Draft', so the target will remain Draft.",
                PlanningPolicy = new PagePlanningOptions
                {
                    TargetPageServerRelativeUrl = "/sites/target/Pages/source.aspx"
                },
                TargetProbe = new PublishingPageTargetSnapshot(),
                LayoutMaterialization = layoutPlan,
                LayoutTargetProbe = layoutProbe,
                LayoutAdmission = layoutAdmission,
                FieldActions = new List<PageFieldAction>
                {
                    new PageFieldAction
                    {
                        SourceInternalName = "OOCLReference",
                        TargetInternalName = "OOCLReference",
                        Disposition = PageFieldDisposition.EvidenceOnly,
                        Reason = "The field is retained for a future mapper."
                    }
                },
                ExpectedPublishingPageContentSha256 = snapshot.PublishingPageContentSha256,
                RuntimeVerification = new RuntimeVerificationManifest
                {
                    Requirements = new List<RuntimeVerificationRequirement>
                    {
                        new RuntimeVerificationRequirement
                        {
                            Id = "authored-dom-equality",
                            Kind = RuntimeVerificationRequirementKind.AuthoredDomEquality,
                            Required = true,
                            Description = "Normalized authored DOM is equal."
                        }
                    }
                }
            };
            return new PublishingPageMigrationPackage
            {
                PlannedAtUtc = DateTimeOffset.UtcNow,
                ExportedAtUtc = DateTimeOffset.UtcNow.AddMinutes(-1),
                State = PublishingPagePackageState.ApprovalReady,
                Snapshot = snapshot,
                Plan = plan,
                SnapshotDigest = snapshotDigest,
                PlanDigest = PublishingPageDigest.ComputePlanDigest(plan),
                Report = new PublishingPageMigrationReport
                {
                    Summary = "Test report"
                }
            };
        }
    }
}
