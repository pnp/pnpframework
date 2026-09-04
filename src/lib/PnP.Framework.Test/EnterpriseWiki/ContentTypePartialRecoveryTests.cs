using Microsoft.VisualStudio.TestTools.UnitTesting;
using PnP.Framework.Migration.Evidence;
using PnP.Framework.Migration.Schema.ContentTypes;
using PnP.Framework.Migration.Schema.Fields;
using System;
using System.Collections.Generic;
using System.Linq;

namespace PnP.Framework.Test.EnterpriseWiki
{
    [TestClass]
    public class ContentTypePartialRecoveryTests
    {
        [TestMethod]
        public void ExactCreatePlanContentTypeWithMissingLinkIsReconciled()
        {
            var fieldId = Guid.Parse("11111111-1111-1111-1111-111111111111");
            var plan = Plan(fieldId, ContentTypeMaterializationDisposition.CreateOwned);
            var probe = Probe(plan, fieldId, false);
            probe.ExistingDescription = "Inherited parent description";

            var result = ContentTypeTargetAdmissionEvaluator.Evaluate(plan, probe);

            Assert.IsTrue(result.IsEligible);
            Assert.AreEqual(ContentTypeMaterializationDisposition.CreateOwned, result.Disposition);
            Assert.IsTrue(result.Warnings.Any(value => value.Contains("interrupted create")));
        }

        [TestMethod]
        public void RuntimeOnlyContentTypeStillRejectsMissingLink()
        {
            var fieldId = Guid.Parse("11111111-1111-1111-1111-111111111111");
            var plan = Plan(fieldId, ContentTypeMaterializationDisposition.ReuseOwned);
            var probe = Probe(plan, fieldId, false);

            var result = ContentTypeTargetAdmissionEvaluator.Evaluate(plan, probe);

            Assert.IsFalse(result.IsEligible);
        }

        private static ContentTypeMaterializationPlan Plan(
            Guid fieldId,
            ContentTypeMaterializationDisposition disposition)
        {
            const string schema = "<Field Type=\"Text\" ID=\"{11111111-1111-1111-1111-111111111111}\" Name=\"OwnedField\" />";
            return new ContentTypeMaterializationPlan
            {
                Disposition = disposition,
                ContentTypeId = "0x010100AA00112233445566778899AABBCCDDEEFF00",
                Name = "Owned content type",
                Description = string.Empty,
                Group = "Migration",
                ParentContentTypeId = "0x010100AA",
                ParentContentTypeName = "Parent",
                RequiredFieldLinks = new List<ContentTypeFieldLinkSnapshot>
                {
                    new ContentTypeFieldLinkSnapshot
                    {
                        FieldId = fieldId,
                        Name = "OwnedField",
                        Role = FieldSchemaRole.DirectBinding
                    }
                },
                Fields = new List<FieldSchemaMaterializationPlan>
                {
                    new FieldSchemaMaterializationPlan
                    {
                        FieldId = fieldId,
                        InternalName = "OwnedField",
                        TypeAsString = "Text",
                        Role = FieldSchemaRole.DirectBinding,
                        Ownership = FieldOwnership.UserDefined,
                        Disposition = FieldSchemaMaterializationDisposition.CreateOrReuseOwned,
                        TargetSchemaXml = schema,
                        TargetPortableSchemaSha256 = FieldSchemaCanonicalizer.PortableDigest(schema)
                    }
                }
            };
        }

        private static ContentTypeTargetProbe Probe(
            ContentTypeMaterializationPlan plan,
            Guid fieldId,
            bool includeLink)
        {
            return new ContentTypeTargetProbe
            {
                Availability = EvidenceAvailability.Captured,
                CanManageContentTypes = true,
                ParentContentTypeAvailable = true,
                ResolvedParentContentTypeId = plan.ParentContentTypeId,
                ContentTypeExists = true,
                ExistingName = plan.Name,
                ExistingDescription = plan.Description,
                ExistingGroup = plan.Group,
                ExistingParentContentTypeId = plan.ParentContentTypeId,
                ExistingFieldLinks = includeLink
                    ? new List<ContentTypeFieldLinkTargetProbe>
                    {
                        new ContentTypeFieldLinkTargetProbe { FieldId = fieldId }
                    }
                    : new List<ContentTypeFieldLinkTargetProbe>(),
                Fields = new List<FieldSchemaTargetProbe>
                {
                    new FieldSchemaTargetProbe
                    {
                        FieldId = fieldId,
                        Exists = true,
                        InternalName = "OwnedField",
                        TypeAsString = "Text",
                        PortableSchemaSha256 = plan.Fields[0].TargetPortableSchemaSha256
                    }
                }
            };
        }
    }
}
