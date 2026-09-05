using Microsoft.VisualStudio.TestTools.UnitTesting;
using PnP.Framework.Migration.Lists.Capture;
using PnP.Framework.Migration.Lists.Fields;
using System;

namespace PnP.Framework.Test.EnterpriseWiki
{
    [TestClass]
    public class ListDependencyClosurePolicyTests
    {
        [TestMethod]
        public void TaxonomyHiddenListIsNotExpandedAsLookupDependency()
        {
            var field = new ListFieldSnapshot
            {
                TypeAsString = "TaxonomyFieldTypeMulti",
                SourceLookupListId = Guid.Parse("11111111-1111-1111-1111-111111111111")
            };

            Assert.IsFalse(ListDependencyClosureSnapshotReader.ShouldFollowLookupDependency(field));
        }

        [TestMethod]
        public void OrdinaryLookupListRemainsInDependencyClosure()
        {
            var field = new ListFieldSnapshot
            {
                TypeAsString = "LookupMulti",
                SourceLookupListId = Guid.Parse("11111111-1111-1111-1111-111111111111")
            };

            Assert.IsTrue(ListDependencyClosureSnapshotReader.ShouldFollowLookupDependency(field));
        }

        [TestMethod]
        public void TaxonomyCatchAllCacheDoesNotExpandHiddenList()
        {
            var field = new ListFieldSnapshot
            {
                Id = Guid.Parse("f3b0adf9-c1a2-4b02-920d-943fba4b3611"),
                InternalName = "TaxCatchAll",
                TypeAsString = "LookupMulti",
                Hidden = true,
                Sealed = true,
                SourceLookupListId = Guid.Parse("11111111-1111-1111-1111-111111111111")
            };

            Assert.IsFalse(ListDependencyClosureSnapshotReader.ShouldFollowLookupDependency(field));
            Assert.IsTrue(ReviewedListRuntimeFieldCatalog.IsSnapshotOnly(field));
        }

        [TestMethod]
        public void UnifiedComplianceMetadataIsSnapshotOnly()
        {
            var field = new ListFieldSnapshot
            {
                Id = Guid.Parse("97dfa283-6ba8-4e68-ab6f-846e53f6d381"),
                InternalName = "_ip_UnifiedCompliancePolicyProperties",
                TypeAsString = "Note"
            };

            Assert.IsTrue(ReviewedListRuntimeFieldCatalog.IsSnapshotOnly(field));
        }
    }
}
