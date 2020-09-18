using Microsoft.VisualStudio.TestTools.UnitTesting;
using PnP.Framework.Test.Framework.Functional.Implementation;

namespace PnP.Framework.Test.Framework.Functional
{
    /// <summary>
    /// Summary description for NavigationTests
    /// </summary>
    [TestClass]
    public class NavigationTests : FunctionalTestBase
    {
        #region construction
        public NavigationTests()
        {
            //debugMode = true;
            //centralSiteCollectionUrl = "https://bertonline.sharepoint.com/sites/TestPnPSC_12345_fdde9ab9-339a-443c-a951-7d1f30cf582b";
            //centralSubSiteUrl = "https://bertonline.sharepoint.com/sites/TestPnPSC_12345_fdde9ab9-339a-443c-a951-7d1f30cf582b/sub";
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

        [TestInitialize()]
        public override void Initialize()
        {
            base.Initialize();

            if (TestCommon.AppOnlyTesting())
            {
                Assert.Inconclusive("Test that require taxonomy creation are not supported in app-only.");
            }
        }
        #endregion

        #region Site collection test cases
        /// <summary>
        /// Navigation test
        /// </summary>
        [TestMethod]
        [Timeout(15 * 60 * 1000)]
        public void SiteCollectionNavigationTest()
        {
            new NavigationImplementation().SiteCollectionNavigation(centralSiteCollectionUrl);
        }

        #endregion

        #region WebTest
        [TestMethod]
        [Timeout(15 * 60 * 1000)]
        public void WebNavigationTest()
        {
            new NavigationImplementation().WebNavigation(centralSubSiteUrl);
        }
        #endregion


    }
}
