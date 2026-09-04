using Microsoft.VisualStudio.TestTools.UnitTesting;
using PnP.Framework.Migration.Evidence;
using PnP.Framework.Migration.Features;
using PnP.Framework.Migration.Packaging;
using PnP.Framework.Migration.Pages.Ingredients;
using PnP.Framework.Migration.Pages.Publishing.Capture;
using PnP.Framework.Migration.Pages.Publishing.Ingredients;
using PnP.Framework.Migration.Pages.Publishing.Layouts;
using PnP.Framework.Migration.Pages.Publishing.Planning;
using PnP.Framework.Migration.Pages.Publishing.Profiles;
using PnP.Framework.Migration.Schema.ContentTypes;
using PnP.Framework.Migration.Schema.Fields;
using PnP.Framework.Migration.Taxonomy;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace PnP.Framework.Test.EnterpriseWiki
{
    [TestClass]
    public class PublishingPageLayoutTargetRuntimeSchemaTests
    {
        [TestMethod]
        public void PartialRuntimeClosureProducesExactReuseOnlyPlan()
        {
            var schema = CreatePartialRuntimeSchema();

            var created = ContentTypeSchemaPlanner.TryCreateTargetRuntimeRequirement(schema, out var plan);

            Assert.IsTrue(created);
            Assert.IsNotNull(plan);
            Assert.AreEqual(ContentTypeMaterializationDisposition.ReuseOwned, plan.Disposition);
            Assert.IsTrue(plan.Fields.All(value =>
                value.Ownership == FieldOwnership.TargetRuntime
                && value.Disposition == FieldSchemaMaterializationDisposition.RequireTargetRuntime
                && value.TargetSchemaXml == null));
            StringAssert.Contains(plan.Reason, "creation is forbidden");
        }

        [TestMethod]
        public void PartialUserDefinedClosureCannotProduceTargetRuntimePlan()
        {
            var schema = CreatePartialRuntimeSchema();
            var field = CreateField(
                Guid.Parse("11111111-1111-1111-1111-111111111111"),
                "CustomerField",
                "Text",
                FieldSchemaRole.DirectBinding,
                "<Field ID=\"{11111111-1111-1111-1111-111111111111}\" Name=\"CustomerField\" Type=\"Text\" />");
            schema.RequiredFieldLinks = new List<ContentTypeFieldLinkSnapshot>
            {
                Link(field)
            };
            schema.RequiredFieldClosure = new List<FieldSchemaSnapshot> { field };

            var created = ContentTypeSchemaPlanner.TryCreateTargetRuntimeRequirement(schema, out var plan);

            Assert.IsFalse(created);
            Assert.IsNull(plan);
        }

        [TestMethod]
        public void TargetRuntimePlanRequiresExistingExactContentType()
        {
            var schema = CreatePartialRuntimeSchema();
            Assert.IsTrue(ContentTypeSchemaPlanner.TryCreateTargetRuntimeRequirement(schema, out var plan));
            var probe = CreateExactTargetProbe(schema);
            probe.ContentTypeExists = false;
            probe.ExistingFieldLinks.Clear();

            var admission = ContentTypeTargetAdmissionEvaluator.Evaluate(plan, probe);

            Assert.IsFalse(admission.IsEligible);
            Assert.AreEqual(ContentTypeMaterializationDisposition.Block, admission.Disposition);
            Assert.IsTrue(admission.Issues.Any(value => value.Code == "TargetRuntimeContentTypeUnavailable"));
        }

        [TestMethod]
        public void TargetRuntimePlanAdmitsExactExistingContentTypeWithoutWrites()
        {
            var schema = CreatePartialRuntimeSchema();
            Assert.IsTrue(ContentTypeSchemaPlanner.TryCreateTargetRuntimeRequirement(schema, out var plan));

            var admission = ContentTypeTargetAdmissionEvaluator.Evaluate(plan, CreateExactTargetProbe(schema));

            Assert.IsTrue(admission.IsEligible);
            Assert.AreEqual(ContentTypeMaterializationDisposition.ReuseOwned, admission.Disposition);
            Assert.IsTrue(admission.Warnings.Any(value => value.Contains("no field or content type schema will be created or repaired")));
        }

        [TestMethod]
        public void TargetRuntimeTaxonomyCompanionAllowsRuntimeInternalName()
        {
            var schema = CreatePartialRuntimeSchema();
            var companion = CreateField(
                Guid.Parse("f863c21f-5fdb-4a91-bb0c-5ae889190dd7"),
                "Wiki_x0020_Page_x0020_CategoriesTaxHTField0",
                "Note",
                FieldSchemaRole.Dependency,
                "<Field ID=\"{f863c21f-5fdb-4a91-bb0c-5ae889190dd7}\" Name=\"Wiki_x0020_Page_x0020_CategoriesTaxHTField0\" Type=\"Note\" Hidden=\"TRUE\" />");
            companion.Hidden = true;
            var taxonomy = CreateField(
                Guid.Parse("e1a5b98c-dd71-426d-acb6-e478c7a5882f"),
                "Wiki_x0020_Page_x0020_Categories",
                "TaxonomyFieldTypeMulti",
                FieldSchemaRole.InheritedFromParent,
                "<Field ID=\"{e1a5b98c-dd71-426d-acb6-e478c7a5882f}\" Name=\"Wiki_x0020_Page_x0020_Categories\" Type=\"TaxonomyFieldTypeMulti\" />");
            taxonomy.Taxonomy = new TaxonomyFieldBindingSnapshot { HiddenTextFieldId = companion.Id };
            schema.RequiredFieldLinks.Add(Link(taxonomy));
            schema.RequiredFieldClosure.Add(companion);
            schema.RequiredFieldClosure.Add(taxonomy);
            Assert.IsTrue(ContentTypeSchemaPlanner.TryCreateTargetRuntimeRequirement(schema, out var plan));
            var probe = CreateExactTargetProbe(schema);
            probe.Fields.Single(value => value.FieldId == companion.Id).InternalName = taxonomy.Id.ToString("N");

            var admission = ContentTypeTargetAdmissionEvaluator.Evaluate(plan, probe);

            Assert.IsTrue(admission.IsEligible);
            Assert.IsTrue(admission.Warnings.Any(value => value.Contains("generated taxonomy companion")));
        }

        [TestMethod]
        public void PartialRuntimeSchemaKeepsLayoutExecutableAndProjectsSubstitution()
        {
            var layout = CreatePartialRuntimeLayout();
            var layoutPlan = PublishingPageLayoutPlanFactory.Create(
                layout,
                new Uri("https://source.sharepoint.com/sites/source"),
                new Uri("https://target.sharepoint.com/sites/target"),
                new Uri("https://target.sharepoint.com/sites/target"),
                "EnterpriseWiki.aspx");

            Assert.AreEqual(PublishingPageLayoutMaterializationDisposition.CreateOwned, layoutPlan.Disposition);
            Assert.AreEqual(ContentTypeMaterializationDisposition.ReuseOwned, layoutPlan.ContentTypeSchema.Disposition);

            var layoutProbe = new PublishingPageLayoutTargetProbe
            {
                TargetServerRelativeUrl = layoutPlan.TargetServerRelativeUrl,
                FileExists = false,
                AssociatedContentTypeAvailable = true,
                ResolvedAssociatedContentTypeId = layoutPlan.ContentTypeSchema.ContentTypeId,
                CanAddAndCustomizePages = true,
                ContentTypeSchema = CreateExactTargetProbe(layout.AssociatedContentTypeSchema),
                Availability = EvidenceAvailability.Captured
            };
            var layoutAdmission = PublishingPageLayoutTargetAdmissionEvaluator.Evaluate(layoutPlan, layoutProbe);
            Assert.IsTrue(layoutAdmission.IsEligible);

            var actions = new Dictionary<string, PageIngredientAction>(StringComparer.Ordinal);
            PublishingPageLayoutIngredientActionProjector.Project(
                new PublishingPageCaptureBundle { Layout = layout },
                new PublishingPageMigrationPlan
                {
                    LayoutMaterialization = layoutPlan,
                    LayoutTargetProbe = layoutProbe,
                    LayoutAdmission = layoutAdmission
                },
                actions);

            var contentType = actions["content-type:page"];
            Assert.AreEqual(IngredientDisposition.Substitute, contentType.Disposition);
            Assert.AreEqual("reuse-exact-target-runtime-content-type", contentType.Realization);
        }

        [TestMethod]
        public void ReviewedStockLayoutProjectsEmbeddedResourcesAsSubstitutions()
        {
            var reference = "~sitecollection/Style Library/~language/Core Styles/page-layouts-21.css";
            var snapshot = new PublishingPageCaptureBundle
            {
                Layout = new PublishingPageLayoutSnapshot
                {
                    ResourceArtifacts = new List<PublishingPageLayoutResourceSnapshot>
                    {
                        new PublishingPageLayoutResourceSnapshot
                        {
                            Reference = new PublishingPageLayoutResourceReference
                            {
                                Attribute = "control:SharePoint:CssRegistration:name",
                                Value = reference
                            },
                            EvidenceState = PublishingPageLayoutResourceEvidenceState.TargetRuntime,
                            ResolvedSourceUrl = reference
                        }
                    }
                }
            };
            var actions = new Dictionary<string, PageIngredientAction>(StringComparer.Ordinal);

            PublishingPageLayoutIngredientActionProjector.Project(
                snapshot,
                new PublishingPageMigrationPlan
                {
                    LayoutMaterialization = new PublishingPageLayoutMaterializationPlan
                    {
                        Disposition = PublishingPageLayoutMaterializationDisposition.ReuseTargetStock,
                        ResourceMaterializations = new List<PublishingPageLayoutResourceMaterializationPlan>()
                    },
                    LayoutAdmission = new PublishingPageLayoutTargetAdmission
                    {
                        IsEligible = true,
                        Disposition = PublishingPageLayoutMaterializationDisposition.ReuseTargetStock
                    }
                },
                actions);

            var resource = actions[PublishingPageIngredientIds.LayoutResource(reference)];
            Assert.AreEqual(IngredientCapability.Available, resource.Capability);
            Assert.AreEqual(IngredientDisposition.Substitute, resource.Disposition);
            Assert.AreEqual("reuse-reviewed-stock-layout-resource", resource.Realization);
            Assert.AreEqual(reference, resource.TargetIdentity);
        }

        [TestMethod]
        public void EnterpriseWikiPolicyAssignsLayoutSystemFieldsToTheirActualOwners()
        {
            Assert.IsTrue(EnterpriseWikiV1WorkflowPolicy.Instance.FieldsHandledByPageWriter.Contains("FileLeafRef"));
            Assert.IsTrue(FieldOwnershipClassifier.IsTargetRuntime(
                Guid.Parse("d31655d1-1d5b-4511-95a1-7a09e9b75bf2"),
                "<Field ID=\"{d31655d1-1d5b-4511-95a1-7a09e9b75bf2}\" Name=\"Editor\" Type=\"User\" SourceID=\"http://schemas.microsoft.com/sharepoint/v3\" />"));
            Assert.IsTrue(FieldOwnershipClassifier.IsTargetRuntime(
                Guid.Parse("8f6b6dd8-9357-4019-8172-966fcd502ed2"),
                "<Field ID=\"{8f6b6dd8-9357-4019-8172-966fcd502ed2}\" Name=\"TaxCatchAllLabel\" Type=\"LookupMulti\" />"));
        }

        [TestMethod]
        public void SourceOwnedCalculatedFieldHasCreateOnlySchemaPlan()
        {
            var calculated = CreateField(
                Guid.Parse("702eb418-d00c-4579-bf9b-f5ac49582083"),
                "Reviews",
                "Calculated",
                FieldSchemaRole.DirectBinding,
                "<Field ID=\"{702eb418-d00c-4579-bf9b-f5ac49582083}\" Name=\"Reviews\" Type=\"Calculated\" ReadOnly=\"TRUE\"><Formula>=1</Formula></Field>");
            calculated.ReadOnly = true;
            var schema = new ContentTypeSchemaSnapshot
            {
                EvidenceState = ContentTypeSchemaEvidenceState.Readable,
                Availability = EvidenceAvailability.Captured,
                SourceWebUrl = "https://source.sharepoint.com/sites/source",
                ContentTypeId = "0x010100AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA",
                Name = "Custom document",
                ParentContentTypeId = BuiltInContentTypeId.Document,
                ParentContentTypeName = "Document",
                RequiredFieldLinks = new List<ContentTypeFieldLinkSnapshot> { Link(calculated) },
                RequiredFieldClosure = new List<FieldSchemaSnapshot> { calculated }
            };

            var plan = ContentTypeSchemaPlanner.CreateRequiredClosure(schema);

            Assert.AreEqual(ContentTypeMaterializationDisposition.CreateOwned, plan.Disposition);
            Assert.AreEqual(
                FieldSchemaMaterializationDisposition.CreateOrReuseOwned,
                plan.Fields.Single().Disposition);
            Assert.IsNotNull(plan.Fields.Single().TargetSchemaXml);
        }

        [TestMethod]
        public void RuntimeCatalogRecognizesLinkToDocumentAndRequiresDocumentIdService()
        {
            var docId = CreateField(
                Guid.Parse("ae3e2a36-125d-45d3-9051-744b513536a6"),
                "_dlc_DocId",
                "Text",
                FieldSchemaRole.InheritedFromParent,
                "<Field ID=\"{ae3e2a36-125d-45d3-9051-744b513536a6}\" Name=\"_dlc_DocId\" Type=\"Text\" />");
            var schema = new ContentTypeSchemaSnapshot
            {
                EvidenceState = ContentTypeSchemaEvidenceState.Readable,
                Availability = EvidenceAvailability.Captured,
                ContentTypeId = "0x010100AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA",
                Name = "Custom document",
                ParentContentTypeId = BuiltInContentTypeId.Document,
                ParentContentTypeName = "Document",
                RequiredFieldLinks = new List<ContentTypeFieldLinkSnapshot>
                {
                    new ContentTypeFieldLinkSnapshot
                    {
                        FieldId = docId.Id,
                        Name = docId.InternalName,
                        Role = docId.Role
                    }
                },
                RequiredFieldClosure = new List<FieldSchemaSnapshot> { docId }
            };

            var features = ContentTypeRuntimeCatalog.CreateFeatureRequirements(
                new[] { BuiltInContentTypeId.LinkToDocument },
                new[] { schema },
                "https://target.sharepoint.com/sites/target");

            Assert.IsTrue(ContentTypeRuntimeCatalog.IsTargetRuntime(BuiltInContentTypeId.LinkToDocument));
            Assert.AreEqual(1, features.Count);
            Assert.AreEqual(ContentTypeRuntimeCatalog.DocumentIdServiceFeatureId, features[0].FeatureId);
            Assert.AreEqual(schema.ContentTypeId, features[0].RequiredByContentTypeIds.Single());

            var plan = ContentTypeSchemaPlanner.CreateRequiredClosure(schema);
            Assert.AreEqual(FieldSchemaMaterializationDisposition.RequireTargetRuntime, plan.Fields.Single().Disposition);
            Assert.IsNotNull(plan.Fields.Single().TargetSchemaXml);
            Assert.AreEqual(
                FieldSchemaCanonicalizer.PortableDigest(docId.SchemaXml),
                plan.Fields.Single().TargetPortableSchemaSha256);
        }

        [TestMethod]
        public void PlanningAdmissionAcceptsSealedParentAndFeaturePredecessors()
        {
            var parentFieldId = Guid.Parse("11111111-1111-1111-1111-111111111111");
            var runtimeFieldId = Guid.Parse("ae3e2a36-125d-45d3-9051-744b513536a6");
            var parentPlan = new ContentTypeMaterializationPlan
            {
                ContentTypeId = "0x010100AA",
                RequiredFieldLinks = new List<ContentTypeFieldLinkSnapshot>
                {
                    new ContentTypeFieldLinkSnapshot { FieldId = parentFieldId, Name = "ParentField" }
                }
            };
            var context = new ContentTypeTargetAdmissionContext(new[] { runtimeFieldId });
            context.RegisterAdmitted(parentPlan);
            var childPlan = new ContentTypeMaterializationPlan
            {
                Disposition = ContentTypeMaterializationDisposition.CreateOwned,
                ContentTypeId = "0x010100AA0011",
                Name = "Child",
                ParentContentTypeId = parentPlan.ContentTypeId,
                ParentContentTypeName = "Parent",
                RequiredFieldLinks = new List<ContentTypeFieldLinkSnapshot>
                {
                    new ContentTypeFieldLinkSnapshot { FieldId = parentFieldId, Name = "ParentField", Role = FieldSchemaRole.InheritedFromParent },
                    new ContentTypeFieldLinkSnapshot { FieldId = runtimeFieldId, Name = "_dlc_DocId", Role = FieldSchemaRole.InheritedFromParent }
                },
                Fields = new List<FieldSchemaMaterializationPlan>
                {
                    new FieldSchemaMaterializationPlan { FieldId = parentFieldId, InternalName = "ParentField", TypeAsString = "Text", Role = FieldSchemaRole.InheritedFromParent, Disposition = FieldSchemaMaterializationDisposition.RequireTargetRuntime },
                    new FieldSchemaMaterializationPlan { FieldId = runtimeFieldId, InternalName = "_dlc_DocId", TypeAsString = "Text", Role = FieldSchemaRole.InheritedFromParent, Disposition = FieldSchemaMaterializationDisposition.RequireTargetRuntime }
                }
            };
            var probe = new ContentTypeTargetProbe
            {
                Availability = EvidenceAvailability.Captured,
                CanManageContentTypes = true
            };

            var strict = ContentTypeTargetAdmissionEvaluator.Evaluate(childPlan, probe);
            var planning = ContentTypeTargetAdmissionEvaluator.Evaluate(childPlan, probe, context);

            Assert.IsFalse(strict.IsEligible);
            Assert.IsTrue(planning.IsEligible);
            Assert.AreEqual(ContentTypeMaterializationDisposition.CreateOwned, planning.Disposition);
            Assert.IsTrue(planning.Warnings.Any(value => value.Contains("strict fresh parent readback")));
            Assert.IsTrue(planning.Warnings.Any(value => value.Contains("platform-feature transaction")));
        }

        [TestMethod]
        public void UnresolvedTaxonomyMappingPreservesMissingTermSetIdentity()
        {
            var sourceStoreId = Guid.Parse("e385fb40-52d4-4fae-9c5b-3e8ff8a5878e");
            var targetStoreId = Guid.Parse("c5e18914-52aa-4047-8ef6-f9654987b925");
            var missingSetId = Guid.Parse("4e691f0e-5ccf-4b99-a3aa-f66b03e98a37");
            var derivedAbsentSetId = Guid.Parse("e318d4cb-6d03-5d29-90c3-7c554147e0f2");
            var field = CreateField(
                Guid.Parse("42387623-5ddb-4764-94ea-e9d826afa77c"),
                "ActivityName",
                "TaxonomyFieldType",
                FieldSchemaRole.DirectBinding,
                "<Field ID=\"{42387623-5ddb-4764-94ea-e9d826afa77c}\" Name=\"ActivityName\" Type=\"TaxonomyFieldType\" />");
            field.Taxonomy = new TaxonomyFieldBindingSnapshot
            {
                SourceTermStoreId = sourceStoreId,
                SourceTermSetId = missingSetId,
                HiddenTextFieldId = Guid.Parse("cfd9f3e8-ce6f-4dc0-a87f-18256c8d4dc3")
            };
            var schema = new ContentTypeSchemaSnapshot
            {
                EvidenceState = ContentTypeSchemaEvidenceState.Readable,
                Availability = EvidenceAvailability.Captured,
                SourceWebUrl = "https://source.sharepoint.com/teams/campusipkits",
                ContentTypeId = "0x010100AA0011",
                Name = "IPKit Guidance",
                ParentContentTypeId = "0x010100AA",
                ParentContentTypeName = "Enterprise Wiki Page",
                RequiredFieldLinks = new List<ContentTypeFieldLinkSnapshot> { Link(field) },
                RequiredFieldClosure = new List<FieldSchemaSnapshot> { field }
            };
            var mapping = new TaxonomyTargetMapping
            {
                SourceTermStoreId = sourceStoreId,
                SourceTermSetId = missingSetId,
                TargetTermStoreId = targetStoreId,
                TargetTermSetId = derivedAbsentSetId,
                Mode = TaxonomyTargetMappingMode.PreserveUnresolvedSourceReference,
                UnresolvedReferenceTargetVerifiedAbsent = true,
                UnresolvedReferenceEvidenceSha256 = new string('a', 64)
            };

            var plan = ContentTypeSchemaPlanner.CreateRequiredClosure(schema, new[] { mapping });
            var fieldPlan = plan.Fields.Single();

            Assert.AreEqual(TaxonomyTargetMappingMode.PreserveUnresolvedSourceReference, fieldPlan.TaxonomyMappingMode);
            Assert.AreEqual(missingSetId, fieldPlan.SourceTermSetId);
            Assert.AreEqual(derivedAbsentSetId, fieldPlan.TargetTermSetId);
            Assert.AreEqual(new string('a', 64), fieldPlan.UnresolvedReferenceEvidenceSha256);
            StringAssert.Contains(fieldPlan.Reason, "source GUID");
        }

        private static PublishingPageLayoutSnapshot CreatePartialRuntimeLayout()
        {
            var bytes = Encoding.UTF8.GetBytes(
                "<%@ Page %><PublishingWebControls:RichHtmlField FieldName=\"PublishingPageContent\" runat=\"server\" /><SharePoint:FormField FieldName=\"Editor\" runat=\"server\" />");
            var schema = CreatePartialRuntimeSchema();
            return new PublishingPageLayoutSnapshot
            {
                EvidenceState = PublishingPageLayoutEvidenceState.Readable,
                Availability = EvidenceAvailability.Captured,
                Url = "https://source.sharepoint.com/sites/source/_catalogs/masterpage/EnterpriseWiki.aspx",
                ServerRelativeUrl = "/sites/source/_catalogs/masterpage/EnterpriseWiki.aspx",
                FileName = "EnterpriseWiki.aspx",
                AssociatedContentTypeName = schema.Name,
                AssociatedContentTypeId = schema.ContentTypeId,
                Bytes = MigrationArtifact.Describe(bytes, "application/vnd.ms-aspx", "EnterpriseWiki.aspx"),
                ContentBase64 = Convert.ToBase64String(bytes),
                Controls = new List<PublishingPageLayoutControl>
                {
                    new PublishingPageLayoutControl { FieldName = "PublishingPageContent" },
                    new PublishingPageLayoutControl { FieldName = "Editor" }
                },
                AssociatedContentTypeSchema = schema
            };
        }

        private static ContentTypeSchemaSnapshot CreatePartialRuntimeSchema()
        {
            var title = CreateField(
                Guid.Parse("fa564e0f-0c70-4ab9-b863-0177e6ddd247"),
                "Title",
                "Text",
                FieldSchemaRole.InheritedFromParent,
                "<Field ID=\"{fa564e0f-0c70-4ab9-b863-0177e6ddd247}\" Name=\"Title\" Type=\"Text\" SourceID=\"http://schemas.microsoft.com/sharepoint/v3\" />");
            var content = CreateField(
                Guid.Parse("f55c4d88-1f2e-4ad9-aaa8-819af4ee7ee8"),
                "PublishingPageContent",
                "HTML",
                FieldSchemaRole.DirectBinding,
                "<Field ID=\"{f55c4d88-1f2e-4ad9-aaa8-819af4ee7ee8}\" Name=\"PublishingPageContent\" Type=\"HTML\" SourceID=\"http://schemas.microsoft.com/sharepoint/v3\" />");
            return new ContentTypeSchemaSnapshot
            {
                EvidenceState = ContentTypeSchemaEvidenceState.Partial,
                Availability = EvidenceAvailability.Partial,
                SourceWebUrl = "https://source.sharepoint.com/sites/source",
                ContentTypeId = "0x010100AA0011",
                Name = "Enterprise Wiki Page",
                Description = "Runtime publishing content type",
                Group = "Publishing Content Types",
                ParentContentTypeId = "0x010100AA",
                ParentContentTypeName = "Page",
                RequiredFieldLinks = new List<ContentTypeFieldLinkSnapshot>
                {
                    Link(title),
                    Link(content)
                },
                RequiredFieldClosure = new List<FieldSchemaSnapshot> { title, content },
                Diagnostics = new List<string>
                {
                    "The Page Layout also references Editor, which is not a direct content type field link."
                }
            };
        }

        private static ContentTypeTargetProbe CreateExactTargetProbe(ContentTypeSchemaSnapshot schema)
        {
            return new ContentTypeTargetProbe
            {
                ContentTypeId = schema.ContentTypeId,
                ParentContentTypeAvailable = true,
                ResolvedParentContentTypeId = schema.ParentContentTypeId,
                ParentFieldLinks = schema.RequiredFieldLinks
                    .Where(value => value.Role == FieldSchemaRole.InheritedFromParent)
                    .Select(TargetLink)
                    .ToList(),
                ContentTypeExists = true,
                ExistingName = schema.Name,
                ExistingDescription = schema.Description,
                ExistingGroup = schema.Group,
                ExistingReadOnly = schema.ReadOnly,
                ExistingSealed = schema.Sealed,
                ExistingHidden = schema.Hidden,
                ExistingParentContentTypeId = schema.ParentContentTypeId,
                ExistingFieldLinks = schema.RequiredFieldLinks.Select(TargetLink).ToList(),
                Fields = schema.RequiredFieldClosure.Select(value => new FieldSchemaTargetProbe
                {
                    FieldId = value.Id,
                    Exists = true,
                    InternalName = value.InternalName,
                    Title = value.Title,
                    TypeAsString = value.TypeAsString,
                    PortableSchemaSha256 = value.PortableSchemaSha256
                }).ToList(),
                Availability = EvidenceAvailability.Captured
            };
        }

        private static FieldSchemaSnapshot CreateField(
            Guid id,
            string internalName,
            string type,
            FieldSchemaRole role,
            string schemaXml)
        {
            return new FieldSchemaSnapshot
            {
                Id = id,
                InternalName = internalName,
                Title = internalName,
                TypeAsString = type,
                Role = role,
                SchemaXml = schemaXml,
                SchemaXmlSha256 = MigrationDigest.ComputeSha256(schemaXml),
                PortableSchemaSha256 = FieldSchemaCanonicalizer.PortableDigest(schemaXml)
            };
        }

        private static ContentTypeFieldLinkSnapshot Link(FieldSchemaSnapshot field)
        {
            return new ContentTypeFieldLinkSnapshot
            {
                FieldId = field.Id,
                Name = field.InternalName,
                Role = field.Role
            };
        }

        private static ContentTypeFieldLinkTargetProbe TargetLink(ContentTypeFieldLinkSnapshot link)
        {
            return new ContentTypeFieldLinkTargetProbe
            {
                FieldId = link.FieldId,
                Name = link.Name,
                Required = link.Required,
                Hidden = link.Hidden
            };
        }
    }
}
