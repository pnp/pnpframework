using Microsoft.VisualStudio.TestTools.UnitTesting;
using PnP.Framework.Tests.Framework.Functional.Implementation;

namespace PnP.Framework.Tests.Framework.Functional
{
    /// <summary>
    /// Test cases for the provisioning engine Publishing functionality
    /// </summary>
    [TestClass]
    public class LocalizationTest : FunctionalTestBase
    {
        #region Construction
        public LocalizationTest()
        {
            //debugMode = true;
            //centralSiteCollectionUrl = "https://bertonline.sharepoint.com/sites/TestPnPSC_12345_dce8970f-8ed6-408f-8e70-766fcb81cbbe";
            //centralSubSiteUrl = "https://bertonline.sharepoint.com/sites/TestPnPSC_12345_dce8970f-8ed6-408f-8e70-766fcb81cbbe/sub";
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

        /// <summary>
        /// PnPLocalizationTest test
        /// </summary>
        [TestMethod]
        [Timeout(15 * 60 * 1000)]
        public void SiteCollectionsLocalizationTest()
        {
            new LocalizationImplementation().SiteCollectionsLocalization(centralSiteCollectionUrl);
        }
        /// <summary>
        /// PnPLocalizationTest test
        /// </summary>
        [TestMethod]
        [Timeout(15 * 60 * 1000)]
        public void WebLocalizationTest()
        {
            new LocalizationImplementation().WebLocalization(centralSubSiteUrl);
        }

    }
}
