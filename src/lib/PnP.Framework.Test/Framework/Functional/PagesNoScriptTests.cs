using Microsoft.VisualStudio.TestTools.UnitTesting;
using PnP.Framework.Test.Framework.Functional.Implementation;

namespace PnP.Framework.Test.Framework.Functional
{
    [TestClass]
    public class PagesNoScriptTests : FunctionalTestBase
    {
        #region Construction
        public PagesNoScriptTests()
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
        /// PagesTest Test
        /// </summary>
        [TestMethod]
        [Timeout(15 * 60 * 1000)]
        public void SiteCollectionPagesTest()
        {
            new PagesImplementation().SiteCollectionPages(centralSiteCollectionUrl);
        }

        /// <summary>
        /// Client side pages Test
        /// </summary>
        [TestMethod]
        [Timeout(15 * 60 * 1000)]
        public void SiteCollectionClientSidePagesTest()
        {
            new PagesImplementation().SiteCollectionClientSidePages(centralSiteCollectionUrl);
        }
        #endregion

        #region Web test cases
        /// <summary>
        /// PagesTest Test
        /// </summary>
        [TestMethod]
        [Timeout(15 * 60 * 1000)]
        public void WebPagesTest()
        {
            new PagesImplementation().WebPages(centralSubSiteUrl);
        }

        /// <summary>
        /// Client side pages Test
        /// </summary>
        [TestMethod]
        [Timeout(15 * 60 * 1000)]
        public void WebClientSidePagesTest()
        {
            new PagesImplementation().WebClientSidePages(centralSubSiteUrl);
        }
        #endregion
    }
}
