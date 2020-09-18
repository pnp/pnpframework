using Microsoft.VisualStudio.TestTools.UnitTesting;
using PnP.Framework.Tests.Framework.Functional.Implementation;

namespace PnP.Framework.Tests.Framework.Functional
{
    /// <summary>
    /// Test cases for the provisioning engine security functionality
    /// </summary>
    [TestClass]
    public class SecurityTests : FunctionalTestBase
    {
        #region Construction
        public SecurityTests()
        {
            //debugMode = true;
            //centralSiteCollectionUrl = "https://bertonline.sharepoint.com/sites/TestPnPSC_12345_6232f367-56a0-4e76-9208-6204b506d401";
            //centralSubSiteUrl = "https://bertonline.sharepoint.com/sites/TestPnPSC_12345_6232f367-56a0-4e76-9208-6204b506d401/sub";
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
        /// Security Test
        /// </summary>
        [TestMethod]
        [Timeout(15 * 60 * 1000)]
        public void SiteCollectionSecurityTest()
        {
            new SecurityImplementation().SiteCollectionSecurity(centralSiteCollectionUrl);
        }
        #endregion
    }
}
