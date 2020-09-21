using Microsoft.VisualStudio.TestPlatform.ObjectModel;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using PnP.Framework.Test.Framework.Functional.Implementation;

namespace PnP.Framework.Test.Framework.Functional
{
    [TestClass]
    public class FilesTests : FunctionalTestBase
    {
        #region Construction
        public FilesTests()
        {
            //centralSiteCollectionUrl = "https://bertonline.sharepoint.com/sites/TestPnPSC_12345_742b8043-e886-4d54-a62c-c9509cb90993";
            //centralSubSiteUrl = "https://bertonline.sharepoint.com/sites/TestPnPSC_12345_742b8043-e886-4d54-a62c-c9509cb90993/sub";
        }
        #endregion

        #region Test setup
        [ClassInitialize()]
        public static void ClassInit(TestContext context)
        {
            //debugMode = false;
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
        /// FilesTest Test
        /// </summary>
        [TestMethod]
        [Timeout(15 * 60 * 1000)]
        public void SiteCollectionFilesTest()
        {
            new FilesImplementation().SiteCollectionFiles(centralSiteCollectionUrl);
        }

        /// <summary>
        /// Directory Files Test
        /// </summary>
        [TestMethod]
        [Timeout(15 * 60 * 1000)]
        public void SiteCollectionDirectoryFilesTest()
        {
            new FilesImplementation().SiteCollectionDirectoryFiles(centralSiteCollectionUrl);
        }
        #endregion

        #region Web test cases
        /// <summary>
        /// FilesTest Test
        /// </summary>
        [TestMethod]
        [Timeout(15 * 60 * 1000)]
        public void WebFilesTest()
        {
            new FilesImplementation().WebFiles(centralSiteCollectionUrl);
        }

        /// <summary>
        /// Directory Files Test
        /// </summary>
        [TestMethod]
        [Timeout(15 * 60 * 1000)]
        public void WebCollectionDirectoryFilesTest()
        {
            new FilesImplementation().WebDirectoryFiles(centralSiteCollectionUrl);
        }

        #endregion


    }
}
