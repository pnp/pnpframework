using Microsoft.VisualStudio.TestTools.UnitTesting;
using PnP.Framework.Migration.Lists.Fields;
using System;

namespace PnP.Framework.Test.EnterpriseWiki
{
    [TestClass]
    public class ReviewedListRuntimeFieldCatalogTests
    {
        [DataTestMethod]
        [DataRow("8ea9462f-400b-428b-9537-ee4a245bc805", "MediaServiceBillingMetadata", "Note")]
        [DataRow("b887b6b2-4dcf-34fc-98b1-d5a42c605755", "MediaServiceFastMetadata", "Note")]
        [DataRow("617f8947-74b2-36bc-9f7e-21ded7029bb5", "MediaServiceMetadata", "Note")]
        [DataRow("6e9ab27b-184d-411f-aedb-00386b5f89c2", "MediaServiceObjectDetectorVersions", "Text")]
        [DataRow("dc412798-e1ac-4884-afa6-3cb4a10dcdda", "MediaServiceSearchProperties", "Note")]
        [DataRow("d3c9caf7-044c-4c71-ae64-092981e54b33", "SharedWithDetails", "Note")]
        [DataRow("ef991a83-108d-4407-8ee5-ccc0c3d836b9", "SharedWithUsers", "UserMulti")]
        [DataRow("89a0fbbc-67b4-40bd-8c9a-386b325ab4ca", "SharingHintHash", "Text")]
        public void ExactReviewedIdentityIsSnapshotOnly(string id, string internalName, string typeAsString)
        {
            var field = Field(id, internalName, typeAsString);

            Assert.IsTrue(ReviewedListRuntimeFieldCatalog.IsSnapshotOnly(field));
        }

        [DataTestMethod]
        [DataRow("617f8947-74b2-36bc-9f7e-21ded7029bb5", "MediaServiceMetadataChanged", "Note")]
        [DataRow("617f8947-74b2-36bc-9f7e-21ded7029bb5", "MediaServiceMetadata", "Text")]
        [DataRow("11111111-2222-3333-4444-555555555555", "MediaServiceMetadata", "Note")]
        public void CatalogRequiresExactIdNameAndType(string id, string internalName, string typeAsString)
        {
            var field = Field(id, internalName, typeAsString);

            Assert.IsFalse(ReviewedListRuntimeFieldCatalog.IsSnapshotOnly(field));
        }

        [TestMethod]
        public void CatalogRequiresReadOnlyAndSealed()
        {
            var field = Field("d3c9caf7-044c-4c71-ae64-092981e54b33", "SharedWithDetails", "Note");
            field.Sealed = false;

            Assert.IsFalse(ReviewedListRuntimeFieldCatalog.IsSnapshotOnly(field));

            field.Sealed = true;
            field.ReadOnly = false;

            Assert.IsFalse(ReviewedListRuntimeFieldCatalog.IsSnapshotOnly(field));
        }

        private static ListFieldSnapshot Field(string id, string internalName, string typeAsString)
        {
            return new ListFieldSnapshot
            {
                Id = Guid.Parse(id),
                InternalName = internalName,
                TypeAsString = typeAsString,
                ReadOnly = true,
                Sealed = true
            };
        }
    }
}
