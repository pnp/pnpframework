using Microsoft.VisualStudio.TestTools.UnitTesting;
using PnP.Framework.Test.Framework.Functional.Implementation;

namespace PnP.Framework.Test.Framework.Functional
{
    [TestClass]
    public class PropertyBagTests : FunctionalTestBase
    {

        #region Construction
        public PropertyBagTests()
        {
            //debugMode = true;
            //centralSiteCollectionUrl = "https://bertonline.sharepoint.com/sites/TestPnPSC_12345_25b60217-025d-45a8-961c-7436cb7419df";
            //centralSubSiteUrl = "https://bertonline.sharepoint.com/sites/TestPnPSC_12345_25b60217-025d-45a8-961c-7436cb7419df/sub";
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
        public void SiteCollectionPropertyBagAddingTest()
        {
            new PropertyBagImplementation().SiteCollectionPropertyBagAdding(centralSiteCollectionUrl);
        }
        #endregion

        #region Web test cases
        [TestMethod]
        [Timeout(15 * 60 * 1000)]
        public void WebPropertyBagAddingTest()
        {
            new PropertyBagImplementation().WebPropertyBagAdding(centralSubSiteUrl);
        }
        #endregion
    }
}
