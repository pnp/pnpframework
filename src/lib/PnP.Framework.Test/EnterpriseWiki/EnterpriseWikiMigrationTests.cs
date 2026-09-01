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
using PnP.Framework.Migration.Pages.ClassicWebParts.Bindings;
using PnP.Framework.Migration.Pages.Publishing.Layouts;
using PnP.Framework.Migration.Lists.Planning;
using PnP.Framework.Migration.Lists.Items;
using PnP.Framework.Migration.Lists.Capture;
using PnP.Framework.Migration.Lists.Fields;
using PnP.Framework.Migration.Lists.ContentTypes;
using PnP.Framework.Migration.Topology;
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
using System.Xml.Linq;

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
            StringAssert.Contains(report, "Site collection and Web topology");
            StringAssert.Contains(report, "List dependency closure");
            StringAssert.Contains(report, "Web Part plan actions");
        }

        [TestMethod]
        public void WebPartPortabilityDelegatesListBindingsAndBlocksKnownUnsupportedTypes()
        {
            const string listView = @"<webParts><webPart xmlns=""http://schemas.microsoft.com/WebPart/v3""><metaData><type name=""Microsoft.SharePoint.WebPartPages.XsltListViewWebPart"" /></metaData><data><properties><property name=""ListId"">58a84d5d-b1ee-4da0-a49b-7e597ee8ae35</property></properties></data></webPart></webParts>";
            const string rss = @"<webParts><webPart xmlns=""http://schemas.microsoft.com/WebPart/v3""><metaData><type name=""Microsoft.SharePoint.Portal.WebControls.RSSAggregatorWebPart"" /></metaData></webPart></webParts>";
            const string scriptEditor = @"<webParts><webPart xmlns=""http://schemas.microsoft.com/WebPart/v3""><metaData><type name=""Microsoft.SharePoint.WebPartPages.ScriptEditorWebPart"" /></metaData></webPart></webParts>";

            Assert.IsNull(EnterpriseWikiWebPartPolicy.GetBlocker(listView));
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

        [TestMethod]
        public void ListWebPartBindingRewritesListWebViewAndPageIdentities()
        {
            var sourceWeb = Guid.Parse("11111111-1111-1111-1111-111111111111");
            var sourceList = Guid.Parse("22222222-2222-2222-2222-222222222222");
            var sourceView = Guid.Parse("33333333-3333-3333-3333-333333333333");
            var targetWeb = Guid.Parse("aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa");
            var targetList = Guid.Parse("bbbbbbbb-bbbb-bbbb-bbbb-bbbbbbbbbbbb");
            var targetView = Guid.Parse("cccccccc-cccc-cccc-cccc-cccccccccccc");
            var exportXml = "<webParts><webPart xmlns=\"http://schemas.microsoft.com/WebPart/v3\"><metaData><type name=\"Microsoft.SharePoint.WebPartPages.XsltListViewWebPart\" /></metaData><data><properties>"
                + "<property name=\"ListId\">" + sourceList.ToString("D") + "</property>"
                + "<property name=\"ListName\">{" + sourceList.ToString("D") + "}</property>"
                + "<property name=\"WebId\">00000000-0000-0000-0000-000000000000</property>"
                + "<property name=\"ViewGuid\">" + sourceView.ToString("D") + "</property>"
                + "<property name=\"TitleUrl\">/teams/source/child/Lists/Resources</property>"
                + "<property name=\"XmlDefinition\">&lt;View Name=\"{" + sourceView.ToString("D") + "}\" Url=\"/teams/source/child/Pages/A.aspx\"&gt;&lt;JSLink&gt;clienttemplates.js&lt;/JSLink&gt;&lt;/View&gt;</property>"
                + "</properties></data></webPart></webParts>";
            var snapshot = new ClassicWebPartSnapshot
            {
                Id = Guid.Parse("44444444-4444-4444-4444-444444444444"),
                Title = "Resources",
                TypeName = "Microsoft.SharePoint.WebPartPages.XsltListViewWebPart",
                ExportXml = exportXml,
                ExportSha256 = MigrationDigest.ComputeSha256(exportXml)
            };

            var parsed = ClassicListWebPartBindingParser.Parse(
                snapshot,
                sourceWeb,
                "https://source.sharepoint.com/teams/source/child",
                "/teams/source/child/Pages/A.aspx");

            Assert.IsTrue(parsed.IsExecutable, string.Join(Environment.NewLine, parsed.Issues.Select(value => value.Message)));
            Assert.AreEqual(sourceWeb, parsed.Binding.SourceListWebId);
            Assert.AreEqual(sourceView, parsed.Binding.SourceViewId);
            Assert.AreEqual("clienttemplates.js", parsed.Binding.JsLink);

            var rewritten = ClassicListWebPartRewriter.Rewrite(parsed.Binding, new ClassicListWebPartTargetMap
            {
                SourceWebId = sourceWeb,
                SourceListId = sourceList,
                SourceViewId = sourceView,
                TargetWebId = targetWeb,
                TargetListId = targetList,
                TargetViewId = targetView,
                TargetListServerRelativeUrl = "/sites/target/child/Lists/Resources",
                TargetListAbsoluteUrl = "https://target.sharepoint.com/sites/target/child/Lists/Resources",
                TargetPageServerRelativeUrl = "/sites/target/child/Pages/A.aspx"
            });
            var properties = XDocument.Parse(rewritten.ExportXml).Descendants()
                .Where(value => value.Name.LocalName == "property")
                .ToDictionary(value => (string)value.Attribute("name"), value => value.Value, StringComparer.OrdinalIgnoreCase);

            Assert.AreEqual(targetWeb.ToString("D"), properties["WebId"]);
            Assert.AreEqual(targetList.ToString("D"), properties["ListId"]);
            Assert.AreEqual("{" + targetList.ToString("D") + "}", properties["ListName"]);
            Assert.AreEqual(targetView.ToString("D"), properties["ViewGuid"]);
            var view = XDocument.Parse(properties["XmlDefinition"]).Root;
            Assert.AreEqual("{" + targetView.ToString("D") + "}", (string)view.Attribute("Name"));
            Assert.AreEqual("/sites/target/child/Pages/A.aspx", (string)view.Attribute("Url"));
        }

        [TestMethod]
        public void LookupDependencyGraphOrdersLookupListsAndBlocksCycles()
        {
            var owner = Guid.Parse("11111111-1111-1111-1111-111111111111");
            var lookup = Guid.Parse("22222222-2222-2222-2222-222222222222");
            var ordered = ListLookupDependencyGraph.Order(
                new[] { owner, lookup },
                new[]
                {
                    new ListLookupDependency
                    {
                        SourceListId = owner,
                        LookupListId = lookup,
                        FieldId = Guid.NewGuid(),
                        FieldInternalName = "Lookup"
                    }
                });

            Assert.IsTrue(ordered.IsExecutable);
            CollectionAssert.AreEqual(new[] { lookup, owner }, ordered.OrderedSourceListIds.ToArray());

            var cycle = ListLookupDependencyGraph.Order(
                new[] { owner, lookup },
                new[]
                {
                    new ListLookupDependency { SourceListId = owner, LookupListId = lookup },
                    new ListLookupDependency { SourceListId = lookup, LookupListId = owner }
                });
            Assert.IsFalse(cycle.IsExecutable);
            Assert.IsTrue(cycle.Issues.Any(value => value.Code == "LookupDependencyCycle"));
        }

        [TestMethod]
        public void TopologyPlannerPreservesNestedWebOwnershipAndStableDigest()
        {
            var siteId = Guid.Parse("11111111-1111-1111-1111-111111111111");
            var rootId = Guid.Parse("22222222-2222-2222-2222-222222222222");
            var childId = Guid.Parse("33333333-3333-3333-3333-333333333333");
            var source = new SourceSiteCollectionSnapshot
            {
                SiteId = siteId,
                SiteCollectionUrl = "https://source.sharepoint.com/teams/source",
                ServerRelativeUrl = "/teams/source",
                RootWebId = rootId,
                Webs = new List<SourceWebSnapshot>
                {
                    new SourceWebSnapshot
                    {
                        SiteId = siteId,
                        WebId = childId,
                        ParentWebId = rootId,
                        SiteCollectionUrl = "https://source.sharepoint.com/teams/source",
                        WebUrl = "https://source.sharepoint.com/teams/source/child",
                        ServerRelativeUrl = "/teams/source/child",
                        Title = "Child",
                        WebTemplate = "CMSPUBLISHING"
                    },
                    new SourceWebSnapshot
                    {
                        SiteId = siteId,
                        WebId = rootId,
                        SiteCollectionUrl = "https://source.sharepoint.com/teams/source",
                        WebUrl = "https://source.sharepoint.com/teams/source",
                        ServerRelativeUrl = "/teams/source",
                        Title = "Root",
                        WebTemplate = "CMSPUBLISHING"
                    }
                }
            };
            var target = new TargetSiteCollectionSpec
            {
                SourceSiteId = siteId,
                Mode = TargetSiteMode.ExistingTargetSite,
                TargetSiteUrl = "https://target.sharepoint.com/sites/target",
                ExpectedTargetSiteId = Guid.Parse("aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa"),
                Title = "Target"
            };
            var policy = new TopologyPlanningPolicy
            {
                WebOverrides = new List<TargetWebOverride>
                {
                    new TargetWebOverride { SourceWebId = childId, TargetUrlSegment = "area", TargetTitle = "Area" }
                }
            };

            var first = new TopologyPlanner().Build(new[] { source }, new[] { target }, policy);
            var second = new TopologyPlanner().Build(new[] { source }, new[] { target }, policy);

            Assert.IsTrue(first.IsExecutable, string.Join(Environment.NewLine, first.Issues.Select(value => value.Message)));
            var child = first.Plan.SiteCollections.Single().Webs.Single(value => value.SourceWebId == childId);
            Assert.AreEqual("/sites/target/area", child.TargetServerRelativeUrl);
            Assert.AreEqual("/sites/target/area/Lists/Resources", TopologyPlanner.MapWebOwnedServerRelativePath(
                "/teams/source/child/Lists/Resources",
                "/teams/source/child",
                child.TargetServerRelativeUrl));
            Assert.AreEqual(first.Plan.PlanDigest, second.Plan.PlanDigest);
        }

        [TestMethod]
        public void ListItemValueCaptureKeepsUnsupportedRawEvidenceForFutureRecovery()
        {
            var captured = ListItemValueSerializer.Serialize("FutureField", new Dictionary<string, object>
            {
                ["reference"] = "OOCL-42",
                ["sequence"] = 7
            });

            Assert.AreEqual(ListItemValueKind.Unsupported, captured.Kind);
            Assert.AreEqual(EvidenceAvailability.Partial, captured.Availability);
            Assert.AreEqual(typeof(Dictionary<string, object>).FullName, captured.RawType);
            StringAssert.Contains(captured.RawValueJson, "OOCL-42");
            Assert.IsTrue(captured.Diagnostics.Any(value => value.Contains("No typed list-item serializer")));
        }

        [TestMethod]
        public void ListPlannerOrdersLookupClosureAndRetainsUnusedUnknownFieldsAsEvidence()
        {
            var siteId = Guid.Parse("11111111-1111-1111-1111-111111111111");
            var webId = Guid.Parse("22222222-2222-2222-2222-222222222222");
            var ownerId = Guid.Parse("33333333-3333-3333-3333-333333333333");
            var lookupId = Guid.Parse("44444444-4444-4444-4444-444444444444");
            var lookupFieldId = Guid.Parse("55555555-5555-5555-5555-555555555555");
            var unknownFieldId = Guid.Parse("66666666-6666-6666-6666-666666666666");
            var owner = CreateListSnapshot(siteId, webId, ownerId, "Owner");
            owner.Fields.Add(new ListFieldSnapshot
            {
                Id = lookupFieldId,
                InternalName = "Category",
                TypeAsString = "Lookup",
                SchemaXml = "<Field ID='{55555555-5555-5555-5555-555555555555}' Name='Category' Type='Lookup' List='{44444444-4444-4444-4444-444444444444}' ShowField='Title' />",
                SourceLookupWebId = webId,
                SourceLookupListId = lookupId,
                LookupField = "Title"
            });
            owner.Fields.Add(new ListFieldSnapshot
            {
                Id = unknownFieldId,
                InternalName = "FutureValue",
                TypeAsString = "FutureType",
                SchemaXml = "<Field ID='{66666666-6666-6666-6666-666666666666}' Name='FutureValue' Type='FutureType' />"
            });
            var lookup = CreateListSnapshot(siteId, webId, lookupId, "Lookup");

            var plan = ListMigrationPlanFactory.Create(
                new[] { owner, lookup },
                new[] { new ListLookupDependency { SourceListId = ownerId, LookupListId = lookupId, FieldId = lookupFieldId, FieldInternalName = "Category" } },
                CreateTopology(siteId, webId),
                null,
                null);

            Assert.IsTrue(plan.IsExecutable, string.Join(Environment.NewLine, plan.Issues.Select(value => value.Message)));
            CollectionAssert.AreEqual(new[] { lookupId, ownerId }, plan.OrderedSourceListIds.ToArray());
            var ownerPlan = plan.Lists.Single(value => value.SourceListId == ownerId);
            Assert.AreEqual(ListFieldMaterializationDisposition.MapLookup, ownerPlan.Fields.Single(value => value.SourceFieldId == lookupFieldId).Disposition);
            Assert.AreEqual(ListFieldMaterializationDisposition.EvidenceOnly, ownerPlan.Fields.Single(value => value.SourceFieldId == unknownFieldId).Disposition);
        }

        [TestMethod]
        public void ListPlannerBlocksNonemptyPrincipalValuesWithoutExplicitMapping()
        {
            var siteId = Guid.Parse("11111111-1111-1111-1111-111111111111");
            var webId = Guid.Parse("22222222-2222-2222-2222-222222222222");
            var listId = Guid.Parse("33333333-3333-3333-3333-333333333333");
            var fieldId = Guid.Parse("44444444-4444-4444-4444-444444444444");
            var source = CreateListSnapshot(siteId, webId, listId, "People");
            source.Fields.Add(new ListFieldSnapshot
            {
                Id = fieldId,
                InternalName = "Owner",
                TypeAsString = "User",
                SchemaXml = "<Field ID='{44444444-4444-4444-4444-444444444444}' Name='Owner' Type='User' />"
            });
            source.SourceItemCount = 1;
            source.Items.Add(new ListItemSnapshot
            {
                SourceItemId = 1,
                Values = new List<ListItemValueSnapshot>
                {
                    new ListItemValueSnapshot { InternalName = "Owner", Kind = ListItemValueKind.User, ScalarValue = "i:0#.f|membership|owner@example.com" }
                }
            });

            var plan = ListMigrationPlanFactory.Create(new[] { source }, null, CreateTopology(siteId, webId), null, null);

            Assert.IsFalse(plan.IsExecutable);
            Assert.AreEqual(ListFieldMaterializationDisposition.Block, plan.Lists.Single().Fields.Single(value => value.SourceFieldId == fieldId).Disposition);
            Assert.IsTrue(plan.Lists.Single().Issues.Any(value => value.Code == "PrincipalMappingUnavailable"));
        }

        [TestMethod]
        public void WebPartReplayCompositionUsesMaterializedListAndViewReceipts()
        {
            var sourceWeb = Guid.Parse("11111111-1111-1111-1111-111111111111");
            var sourceList = Guid.Parse("22222222-2222-2222-2222-222222222222");
            var sourceView = Guid.Parse("33333333-3333-3333-3333-333333333333");
            var webPartId = Guid.Parse("44444444-4444-4444-4444-444444444444");
            var targetWeb = Guid.Parse("aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa");
            var targetList = Guid.Parse("bbbbbbbb-bbbb-bbbb-bbbb-bbbbbbbbbbbb");
            var targetView = Guid.Parse("cccccccc-cccc-cccc-cccc-cccccccccccc");
            var xml = "<webParts><webPart><data><properties>"
                + "<property name='ListId'>" + sourceList.ToString("D") + "</property>"
                + "<property name='ListName'>{" + sourceList.ToString("D") + "}</property>"
                + "<property name='WebId'>" + sourceWeb.ToString("D") + "</property>"
                + "<property name='ViewGuid'>" + sourceView.ToString("D") + "</property>"
                + "<property name='TitleUrl'>/sites/source/Lists/Items</property>"
                + "<property name='XmlDefinition'>&lt;View Name=\"{" + sourceView.ToString("D") + "}\" Url=\"/sites/source/Pages/A.aspx\" /&gt;</property>"
                + "</properties></data></webPart></webParts>";
            var captured = new ClassicWebPartSnapshot { Id = webPartId, ExportXml = xml };
            var binding = new ClassicListWebPartBindingSnapshot
            {
                SourceWebPartId = webPartId,
                SourceListWebId = sourceWeb,
                SourceListId = sourceList,
                SourceViewId = sourceView,
                SourceTitleUrl = "/sites/source/Lists/Items",
                SourceExportXml = xml
            };
            var receipt = new ListMaterializationReceipt
            {
                SourceWebId = sourceWeb,
                SourceListId = sourceList,
                TargetWebId = targetWeb,
                TargetListId = targetList,
                TargetRootFolderServerRelativeUrl = "/sites/target/Lists/Items",
                TargetViewIds = new Dictionary<Guid, Guid> { [sourceView] = targetView }
            };

            var replay = ClassicWebPartReplayComposer.Compose(
                captured,
                new ClassicWebPartAction { SourceWebPartId = webPartId, Disposition = ClassicWebPartDisposition.RebindListAfterMaterialization },
                binding,
                receipt,
                "/sites/target/Pages/A.aspx",
                Array.Empty<PageTextReplacement>());
            var properties = XDocument.Parse(replay).Descendants().Where(value => value.Name.LocalName == "property")
                .ToDictionary(value => (string)value.Attribute("name"), value => value.Value, StringComparer.OrdinalIgnoreCase);

            Assert.AreEqual(targetWeb.ToString("D"), properties["WebId"]);
            Assert.AreEqual(targetList.ToString("D"), properties["ListId"]);
            Assert.AreEqual(targetView.ToString("D"), properties["ViewGuid"]);
            Assert.AreEqual("/sites/target/Lists/Items", properties["TitleUrl"]);
        }

        [TestMethod]
        public void CustomDocumentContentTypeIsCapturedAsClosureInsteadOfMisclassifiedAsRuntime()
        {
            var siteId = Guid.Parse("11111111-1111-1111-1111-111111111111");
            var webId = Guid.Parse("22222222-2222-2222-2222-222222222222");
            var listId = Guid.Parse("33333333-3333-3333-3333-333333333333");
            var fieldId = Guid.Parse("44444444-4444-4444-4444-444444444444");
            const string siteContentTypeId = "0x010100AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA";
            const string listContentTypeId = siteContentTypeId + "00BBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBB";
            var fieldSchema = "<Field ID='{44444444-4444-4444-4444-444444444444}' Name='CaseNumber' DisplayName='Case number' Type='Text' />";
            var source = CreateListSnapshot(siteId, webId, listId, "Documents");
            source.BaseTemplate = 101;
            source.BaseType = "DocumentLibrary";
            source.Fields.Add(new ListFieldSnapshot
            {
                Id = fieldId,
                InternalName = "CaseNumber",
                TypeAsString = "Text",
                SchemaXml = fieldSchema,
                SchemaXmlSha256 = MigrationDigest.ComputeSha256(fieldSchema),
                PortableSchemaSha256 = FieldSchemaCanonicalizer.PortableDigest(fieldSchema)
            });
            source.ContentTypes.Add(new ListContentTypeSnapshot
            {
                Id = listContentTypeId,
                Name = "Case document",
                ParentId = siteContentTypeId,
                FieldLinks = new List<ListContentTypeFieldLinkSnapshot>
                {
                    new ListContentTypeFieldLinkSnapshot { FieldId = fieldId, InternalName = "CaseNumber" }
                }
            });
            source.SiteContentTypes.Add(new ContentTypeSchemaSnapshot
            {
                EvidenceState = ContentTypeSchemaEvidenceState.Readable,
                Availability = EvidenceAvailability.Captured,
                SourceWebUrl = "https://source.sharepoint.com/sites/source",
                SourceScope = "/sites/source",
                ContentTypeId = siteContentTypeId,
                Name = "Case document",
                Hidden = true,
                ReadOnly = true,
                Sealed = true,
                ParentContentTypeId = "0x0101",
                ParentContentTypeName = "Document",
                RequiredFieldLinks = new List<ContentTypeFieldLinkSnapshot>
                {
                    new ContentTypeFieldLinkSnapshot { FieldId = fieldId, Name = "CaseNumber", Role = FieldSchemaRole.DirectBinding }
                },
                RequiredFieldClosure = new List<FieldSchemaSnapshot>
                {
                    new FieldSchemaSnapshot
                    {
                        Id = fieldId,
                        InternalName = "CaseNumber",
                        Title = "Case number",
                        TypeAsString = "Text",
                        SchemaXml = fieldSchema,
                        SchemaXmlSha256 = MigrationDigest.ComputeSha256(fieldSchema),
                        PortableSchemaSha256 = FieldSchemaCanonicalizer.PortableDigest(fieldSchema),
                        Role = FieldSchemaRole.DirectBinding
                    }
                }
            });

            var plan = ListMigrationPlanFactory.Create(new[] { source }, null, CreateTopology(siteId, webId), null, null);
            var listPlan = plan.Lists.Single();

            Assert.IsFalse(ContentTypeRuntimeCatalog.IsTargetRuntime(siteContentTypeId));
            Assert.IsTrue(plan.IsExecutable, string.Join(Environment.NewLine, listPlan.Issues.Select(value => value.Message)));
            Assert.AreEqual(1, listPlan.SiteContentTypes.Count);
            Assert.AreEqual(siteContentTypeId, listPlan.SiteContentTypes[0].Schema.ContentTypeId);
            Assert.IsTrue(listPlan.SiteContentTypes[0].Schema.Hidden);
            Assert.IsTrue(listPlan.SiteContentTypes[0].Schema.ReadOnly);
            Assert.IsTrue(listPlan.SiteContentTypes[0].Schema.Sealed);
            Assert.AreEqual("https://target.sharepoint.com/sites/target", listPlan.SiteContentTypes[0].TargetOwnerWebUrl);
        }

        [TestMethod]
        public void ListSemanticDigestExcludesMutableTargetAnalysis()
        {
            var siteId = Guid.Parse("11111111-1111-1111-1111-111111111111");
            var webId = Guid.Parse("22222222-2222-2222-2222-222222222222");
            var source = CreateListSnapshot(siteId, webId, Guid.Parse("33333333-3333-3333-3333-333333333333"), "Items");
            source.SiteContentTypes.Add(new ContentTypeSchemaSnapshot
            {
                EvidenceState = ContentTypeSchemaEvidenceState.Readable,
                Availability = EvidenceAvailability.Captured,
                SourceWebUrl = "https://source.sharepoint.com/sites/source",
                SourceScope = "/sites/source",
                ContentTypeId = "0x0100AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA",
                Name = "Custom item",
                ParentContentTypeId = "0x01",
                ParentContentTypeName = "Item"
            });
            source.ContentTypes.Add(new ListContentTypeSnapshot
            {
                Id = "0x0100AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA00BBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBB",
                Name = "Custom item",
                ParentId = "0x0100AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
            });
            var planSet = ListMigrationPlanFactory.Create(new[] { source }, null, CreateTopology(siteId, webId), null, null);
            var listPlan = planSet.Lists.Single();
            var before = listPlan.PlanDigest;
            listPlan.SiteContentTypes[0].DeferredUntilTopologyMaterialization = true;
            listPlan.SiteContentTypes[0].TargetProbe = new ContentTypeTargetProbe { ContentTypeId = "changed" };
            listPlan.SiteContentTypes[0].TargetAdmission = new ContentTypeTargetAdmission
            {
                IsEligible = true,
                Disposition = ContentTypeMaterializationDisposition.CreateOwned
            };
            listPlan.TargetProbe = new ListTargetProbe
            {
                TargetWebExists = true,
                Disposition = ListMaterializationDisposition.ReuseOwned
            };
            ListMigrationPlanFactory.SealTargetAnalysis(planSet);

            Assert.AreEqual(ListMaterializationDisposition.ReuseOwned, listPlan.Disposition);
            Assert.AreEqual(before, ListMigrationPlanFactory.ComputePlanDigest(listPlan));
            Assert.AreEqual(planSet.PlanDigest, ListMigrationPlanFactory.ComputeSetDigest(planSet));
        }

        [TestMethod]
        public void ReadOnlyRuntimeListFieldsAreRequiredButTheirValuesAreNotReplayed()
        {
            var siteId = Guid.Parse("11111111-1111-1111-1111-111111111111");
            var webId = Guid.Parse("22222222-2222-2222-2222-222222222222");
            var source = CreateListSnapshot(siteId, webId, Guid.Parse("33333333-3333-3333-3333-333333333333"), "Items");
            const string schema = "<Field ID='{44444444-4444-4444-4444-444444444444}' Name='Modified' DisplayName='Modified' Type='DateTime' ReadOnly='TRUE' SourceID='http://schemas.microsoft.com/sharepoint/v3' />";
            source.Fields.Add(new ListFieldSnapshot
            {
                Id = Guid.Parse("44444444-4444-4444-4444-444444444444"),
                InternalName = "Modified",
                Title = "Modified",
                TypeAsString = "DateTime",
                SchemaXml = schema,
                SchemaXmlSha256 = MigrationDigest.ComputeSha256(schema),
                PortableSchemaSha256 = FieldSchemaCanonicalizer.PortableDigest(schema),
                ReadOnly = true,
                FromBaseType = true
            });
            source.Items.Add(new ListItemSnapshot
            {
                SourceItemId = 1,
                Values = new List<ListItemValueSnapshot>
                {
                    new ListItemValueSnapshot { InternalName = "Modified", Kind = ListItemValueKind.DateTime, ScalarValue = "2026-08-31T00:00:00.0000000Z" }
                }
            });
            source.SourceItemCount = 1;

            var plan = ListMigrationPlanFactory.Create(new[] { source }, null, CreateTopology(siteId, webId), null, null);

            Assert.AreEqual(ListFieldMaterializationDisposition.RequireTargetRuntime, plan.Lists.Single().Fields.Single().Disposition);
        }

        [TestMethod]
        public void RuntimeListFieldCompatibilityPreservesScalarAndCollectionShapes()
        {
            Assert.IsTrue(ListFieldTypeCompatibility.IsCompatibleRuntimeType("Note", "Text"));
            Assert.IsTrue(ListFieldTypeCompatibility.IsCompatibleRuntimeType("Choice", "Text"));
            Assert.IsFalse(ListFieldTypeCompatibility.IsCompatibleRuntimeType("MultiChoice", "Choice"));
            Assert.IsFalse(ListFieldTypeCompatibility.IsCompatibleRuntimeType("UserMulti", "User"));
            Assert.IsFalse(ListFieldTypeCompatibility.IsCompatibleRuntimeType("LookupMulti", "Lookup"));
            Assert.IsFalse(ListFieldTypeCompatibility.IsCompatibleRuntimeType("TaxonomyFieldTypeMulti", "TaxonomyFieldType"));
        }

        [TestMethod]
        public void CalculatedListFieldsArePlannedInFormulaDependencyOrder()
        {
            var siteId = Guid.Parse("11111111-1111-1111-1111-111111111111");
            var webId = Guid.Parse("22222222-2222-2222-2222-222222222222");
            var source = CreateListSnapshot(siteId, webId, Guid.Parse("33333333-3333-3333-3333-333333333333"), "Items");
            var alphaSchema = "<Field ID='{44444444-4444-4444-4444-444444444444}' Name='Alpha' DisplayName='Alpha Result' Type='Calculated' ReadOnly='TRUE'><Formula>=[Zulu Result]+1</Formula></Field>";
            var zuluSchema = "<Field ID='{55555555-5555-5555-5555-555555555555}' Name='Zulu' DisplayName='Zulu Result' Type='Calculated' ReadOnly='TRUE'><Formula>=1</Formula></Field>";
            source.Fields.Add(CalculatedListField(Guid.Parse("44444444-4444-4444-4444-444444444444"), "Alpha", "Alpha Result", alphaSchema));
            source.Fields.Add(CalculatedListField(Guid.Parse("55555555-5555-5555-5555-555555555555"), "Zulu", "Zulu Result", zuluSchema));

            var plan = ListMigrationPlanFactory.Create(new[] { source }, null, CreateTopology(siteId, webId), null, null);
            var fields = plan.Lists.Single().Fields;

            Assert.AreEqual("Zulu", fields[0].InternalName);
            Assert.AreEqual("Alpha", fields[1].InternalName);
        }

        private static ListDependencySnapshot CreateListSnapshot(Guid siteId, Guid webId, Guid listId, string title)
        {
            return new ListDependencySnapshot
            {
                SourceSiteId = siteId,
                SourceWebId = webId,
                SourceWebUrl = "https://source.sharepoint.com/sites/source",
                SourceListId = listId,
                Title = title,
                BaseTemplate = 100,
                BaseType = "GenericList",
                RootFolderServerRelativeUrl = "/sites/source/Lists/" + title,
                Availability = EvidenceAvailability.Captured
            };
        }

        private static ListFieldSnapshot CalculatedListField(Guid id, string internalName, string title, string schema)
        {
            return new ListFieldSnapshot
            {
                Id = id,
                InternalName = internalName,
                Title = title,
                TypeAsString = "Calculated",
                SchemaXml = schema,
                SchemaXmlSha256 = MigrationDigest.ComputeSha256(schema),
                PortableSchemaSha256 = FieldSchemaCanonicalizer.PortableDigest(schema),
                ReadOnly = true
            };
        }

        private static TopologyPlan CreateTopology(Guid siteId, Guid webId)
        {
            var plan = new TopologyPlan
            {
                SiteCollections = new List<SiteCollectionMappingPlan>
                {
                    new SiteCollectionMappingPlan
                    {
                        SourceSiteId = siteId,
                        TargetSiteCollectionUrl = "https://target.sharepoint.com/sites/target",
                        Webs = new List<WebMappingPlan>
                        {
                            new WebMappingPlan
                            {
                                Kind = TopologyNodeKind.SiteCollectionRoot,
                                SourceSiteId = siteId,
                                SourceWebId = webId,
                                SourceServerRelativeUrl = "/sites/source",
                                TargetWebUrl = "https://target.sharepoint.com/sites/target",
                                TargetServerRelativeUrl = "/sites/target"
                            }
                        }
                    }
                }
            };
            plan.PlanDigest = TopologyPlanner.ComputeDigest(plan);
            return plan;
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
