using Microsoft.VisualStudio.TestTools.UnitTesting;
using PnP.Framework.Tests.Framework.Functional.Implementation;

namespace PnP.Framework.Tests.Framework.Functional
{
    /// <summary>
    /// Test cases for the provisioning engine Publishing functionality
    /// </summary>
    [TestClass]
    public class PublishingNoScriptTest : FunctionalTestBase
    {
        #region Construction
        public PublishingNoScriptTest()
        {
            //debugMode = true;
            //centralSiteCollectionUrl = "https://bertonline.sharepoint.com/sites/TestPnPSC_12345_a21016d5-886f-49eb-9984-f9f4dce76741";
            //centralSubSiteUrl = "https://bertonline.sharepoint.com/sites/TestPnPSC_12345_a21016d5-886f-49eb-9984-f9f4dce76741/sub";
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
        /// Design package publishing test in site collection
        /// </summary>
        [TestMethod]
        [Timeout(15 * 60 * 1000)]
        public void SiteCollectionPublishingTest()
        {
            new PublishingImplementation().SiteCollectionPublishing(centralSiteCollectionUrl);
        }
        #endregion        
    }
}
