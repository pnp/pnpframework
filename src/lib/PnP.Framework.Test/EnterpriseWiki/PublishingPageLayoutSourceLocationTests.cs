using Microsoft.VisualStudio.TestTools.UnitTesting;
using PnP.Framework.Migration.Pages.Publishing.Layouts;
using System;

namespace PnP.Framework.Test.EnterpriseWiki
{
    [TestClass]
    public class PublishingPageLayoutSourceLocationTests
    {
        [TestMethod]
        public void ResolvesSameTenantExternalSiteCollectionOwner()
        {
            var resolved = PublishingPageLayoutSourceLocation.TryResolve(
                new Uri("https://source.sharepoint.com/teams/uat"),
                "https://source.sharepoint.com/teams/staging/_catalogs/masterpage/IPKit%20Guidance.aspx",
                out var location,
                out var diagnostic);

            Assert.IsTrue(resolved, diagnostic);
            Assert.AreEqual("https://source.sharepoint.com/teams/staging", location.OwnerSiteCollectionUrl.AbsoluteUri);
            Assert.AreEqual("/teams/staging/_catalogs/masterpage/IPKit Guidance.aspx", location.ServerRelativeUrl);
            Assert.IsTrue(location.IsExternalToPageSiteCollection);
        }

        [TestMethod]
        public void ResolvesPageOwnedServerRelativeLayout()
        {
            var resolved = PublishingPageLayoutSourceLocation.TryResolve(
                new Uri("https://source.sharepoint.com/teams/source"),
                "/teams/source/_catalogs/masterpage/EnterpriseWiki.aspx",
                out var location,
                out var diagnostic);

            Assert.IsTrue(resolved, diagnostic);
            Assert.AreEqual("https://source.sharepoint.com/teams/source", location.OwnerSiteCollectionUrl.AbsoluteUri);
            Assert.IsFalse(location.IsExternalToPageSiteCollection);
        }

        [TestMethod]
        public void RejectsCrossOriginLayoutOwner()
        {
            var resolved = PublishingPageLayoutSourceLocation.TryResolve(
                new Uri("https://source.sharepoint.com/teams/source"),
                "https://other.sharepoint.com/teams/source/_catalogs/masterpage/custom.aspx",
                out _,
                out var diagnostic);

            Assert.IsFalse(resolved);
            StringAssert.Contains(diagnostic, "outside the source tenant HTTPS origin");
        }
    }
}
