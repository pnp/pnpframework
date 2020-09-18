using Microsoft.VisualStudio.TestTools.UnitTesting;
using PnP.Framework.Test.Framework.Functional.Implementation;

namespace PnP.Framework.Test.Framework.Functional
{
    /// <summary>
    /// Test cases for the provisioning engine search settings functionality
    /// </summary>
    [TestClass]
    public class RegionalSettingsNoScriptTests : FunctionalTestBase
    {
        #region Construction
        public RegionalSettingsNoScriptTests()
        {
            //debugMode = true;
            //centralSiteCollectionUrl = "https://bertonline.sharepoint.com/sites/TestPnPSC_12345_23939b71-a69c-4d57-a818-5448a1cdaf28";
            //centralSubSiteUrl = "https://bertonline.sharepoint.com/sites/TestPnPSC_12345_23939b71-a69c-4d57-a818-5448a1cdaf28/sub";
        }
        #endregion

        #region Test setup
        [ClassInitialize()]
        public static void ClassInit(TestContext context)
        {
            ClassInitBase(context, true);
        }

        [ClassCleanup()]
        public static void ClassCleanup()
        {
            ClassCleanupBase();
        }
        #endregion

        #region Site collection test cases
        /// <summary>
        /// Site RegionalSettings Test
        /// </summary>
        [TestMethod]
        [Timeout(15 * 60 * 1000)]
        public void SiteCollectionRegionalSettingsTest()
        {
            new RegionalSettingsImplementation().SiteCollectionRegionalSettings(centralSiteCollectionUrl);
        }
        #endregion

        #region Web test cases
        /// <summary>
        /// Web RegionalSettings test
        /// </summary>
        [TestMethod]
        [Timeout(15 * 60 * 1000)]
        public void WebRegionalSettingsTest()
        {
            new RegionalSettingsImplementation().WebRegionalSettings(centralSubSiteUrl);
        }
        #endregion
    }
}
