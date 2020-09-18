using Microsoft.VisualStudio.TestTools.UnitTesting;
using PnP.Framework.Test.Framework.Functional.Implementation;

namespace PnP.Framework.Test.Framework.Functional
{
    [TestClass]
    public class CustomActionTests : FunctionalTestBase
    {

        #region Construction
        public CustomActionTests()
        {
            //debugMode = true;
            //centralSiteCollectionUrl = "https://bertonline.sharepoint.com/sites/TestPnPSC_12345_ab5f2990-6015-48c5-a09b-685153dcebc9";
            //centralSubSiteUrl = "https://bertonline.sharepoint.com/sites/TestPnPSC_12345_ab5f2990-6015-48c5-a09b-685153dcebc9/sub";
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
        [TestMethod]
        [Timeout(15 * 60 * 1000)]
        public void SiteCollectionCustomActionAddingTest()
        {
            new CustomActionImplementation().SiteCollectionCustomActionAdding(centralSiteCollectionUrl);
        }
        #endregion

        #region Web test cases
        [TestMethod]
        [Timeout(15 * 60 * 1000)]
        public void WebCustomActionAddingTest()
        {
            new CustomActionImplementation().WebCustomActionAdding(centralSubSiteUrl);
        }
        #endregion

    }
}
