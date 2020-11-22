using System;
using Microsoft.SharePoint.Client;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace PnP.Framework.Modernization.Tests.Transform.Utility
{
    [TestClass]
    public class WebExtensionsTest
    {
        [TestMethod]
        public void WebExtensions_GetUrl()
        {
            using (var ctx = TestCommon.CreateClientContext(TestCommon.AppSetting("SPOTargetSiteUrl")))
            {
                var expectedUrl = TestCommon.AppSetting("SPOTargetSiteUrl");
                var result = ctx.Web.GetUrl();

                Assert.AreEqual(expectedUrl, result);

            }
        }
    }
}
