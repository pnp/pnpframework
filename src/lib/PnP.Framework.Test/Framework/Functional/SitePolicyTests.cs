using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace PnP.Framework.Tests.Framework.Functional
{
    /// <summary>
    /// Test cases for the provisioning engine security functionality
    /// </summary>
    [TestClass]
    public class SitePolicyTests : FunctionalTestBase
    {
        #region Construction
        public SitePolicyTests()
        {
            //debugMode = true;
            //centralSiteCollectionUrl = "https://bertonline.sharepoint.com/sites/TestPnPSC_12345_c89c25d3-4153-4464-8ad3-d0d6715fb6a8";
            //centralSubSiteUrl = "https://bertonline.sharepoint.com/sites/TestPnPSC_12345_c89c25d3-4153-4464-8ad3-d0d6715fb6a8/sub";
        }
        #endregion

        #region Test setup
        [ClassInitialize()]
        public static void ClassInit(TestContext context)
        {
            ClassInitBase(context);
        }

        [ClassCleanup()]
        public static void ClassCleanup()
        {
            ClassCleanupBase();
        }
        #endregion

        #region Site collection test cases
        /// <summary>
        /// SitePolicyTests Test
        /// </summary>
        //[TestMethod]
        //[Timeout(15 * 60 * 1000)]
        //public void SitePolicyTest()
        //{
        //    using (var cc = TestCommon.CreateClientContext(centralSiteCollectionUrl))
        //    {
        //        var result = TestProvisioningTemplate(cc, "sitepolicy_add.xml", Handlers.SitePolicy);
        //        SitePolicyValidator spv= new SitePolicyValidator();
        //        Assert.IsTrue(spv.Validate(result.SourceTemplate.SitePolicy, result.TargetTemplate.SitePolicy,result.TargetTokenParser));
        //    }
        //}
        #endregion
    }
}
