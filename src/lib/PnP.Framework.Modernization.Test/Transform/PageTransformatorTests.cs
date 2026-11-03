using Microsoft.SharePoint.Client;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using PnP.Framework.Modernization.Cache;
using PnP.Framework.Modernization.Transform;
using System;

namespace PnP.Framework.Modernization.Tests.Transform
{
    [TestClass]
    public class PageTransformatorTests
    {
        [DataTestMethod]
        [DataRow("https://contoso.sharepoint.com/sites/source", false)]
        [DataRow("https://contoso.sharepoint.com/sites/source/", true)]
        [DataRow("https://contoso.sharepoint.com/sites/source/subweb", true)]
        public void NormalizeSameSiteTargetContext_UsesSourceOnlyModeAndPreservesTargetFolder(string targetUrl, bool initiallyCrossSite)
        {
            using (var sourceContext = new ClientContext("https://contoso.sharepoint.com/sites/source"))
            using (var targetContext = new ClientContext(targetUrl))
            {
                CacheManager.Instance.SetSharePointVersion(new Uri(sourceContext.Url), SPVersion.SPO);
                CacheManager.Instance.SetSharePointVersion(new Uri(targetContext.Url), SPVersion.SPO);
                var transformationInformation = new PageTransformationInformation(null)
                {
                    TargetPageFolder = "Modernized",
                    TargetPageFolderOverridesDefaultFolder = true,
                    TargetPageName = "Converted.aspx"
                };
                var transformator = new PageTransformator(sourceContext, targetContext, new PageTransformation());
                transformator.CrossFarmTransformationValidation(transformationInformation);
                Assert.AreEqual(initiallyCrossSite, transformationInformation.IsCrossSiteTransformation);

                var siteId = Guid.NewGuid();
                var normalized = transformator.NormalizeSameSiteTargetContext(siteId, siteId, transformationInformation);

                Assert.IsTrue(normalized);
                Assert.IsNull(transformator.targetClientContext);
                Assert.AreSame(sourceContext, transformator.sourceClientContext);
                Assert.AreEqual("Modernized", transformationInformation.TargetPageFolder);
                Assert.IsTrue(transformationInformation.TargetPageFolderOverridesDefaultFolder);
                Assert.AreEqual("Converted.aspx", transformationInformation.TargetPageName);
                Assert.IsFalse(transformationInformation.TargetPageTakesSourcePageName);
                Assert.IsFalse(transformationInformation.IsCrossSiteTransformation);
                Assert.IsFalse(transformationInformation.IsCrossFarmTransformation);
                Assert.AreEqual(SPVersion.SPO, transformationInformation.TargetVersion);
                Assert.AreEqual(transformationInformation.SourceVersionNumber, transformationInformation.TargetVersionNumber);
            }
        }

        [TestMethod]
        public void NormalizeSameSiteTargetContext_DifferentSitePreservesTargetSettings()
        {
            using (var sourceContext = new ClientContext("https://contoso.sharepoint.com/sites/source"))
            using (var targetContext = new ClientContext("https://fabrikam.sharepoint.com/sites/target"))
            {
                var information = new PageTransformationInformation(null)
                {
                    IsCrossSiteTransformation = true,
                    IsCrossFarmTransformation = true,
                    SourceVersion = SPVersion.SP2019,
                    TargetVersion = SPVersion.SPO,
                    TargetVersionNumber = "16.0.0.0",
                    TargetPageFolder = "Modernized"
                };
                var transformator = new PageTransformator(sourceContext, targetContext, new PageTransformation());

                Assert.IsFalse(transformator.NormalizeSameSiteTargetContext(Guid.NewGuid(), Guid.NewGuid(), information));
                Assert.AreSame(targetContext, transformator.targetClientContext);
                Assert.IsTrue(information.IsCrossSiteTransformation);
                Assert.IsTrue(information.IsCrossFarmTransformation);
                Assert.AreEqual(SPVersion.SPO, information.TargetVersion);
                Assert.AreEqual("16.0.0.0", information.TargetVersionNumber);
                Assert.AreEqual("Modernized", information.TargetPageFolder);
            }
        }
    }
}
