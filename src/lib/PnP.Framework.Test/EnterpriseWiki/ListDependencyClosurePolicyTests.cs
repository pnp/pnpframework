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
    }
}
