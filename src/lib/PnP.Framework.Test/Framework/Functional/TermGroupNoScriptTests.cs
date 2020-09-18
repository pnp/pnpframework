using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.Taxonomy;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using PnP.Framework.Enums;
using PnP.Framework.Provisioning.Model;
using PnP.Framework.Provisioning.ObjectHandlers;
using PnP.Framework.Provisioning.Providers.Xml;
using PnP.Framework.Tests.Framework.Functional.Implementation;
using PnP.Framework.Tests.Framework.Functional.Validators;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;
using System.Xml.XPath;

namespace PnP.Framework.Tests.Framework.Functional
{
    /// <summary>
    /// Test cases for the provisioning engine term group functionality
    /// </summary>
    [TestClass]
    public class TermGroupNoScriptTests : FunctionalTestBase
    {
        #region Construction
        public TermGroupNoScriptTests()
        {
            //debugMode = true;
            //centralSiteCollectionUrl = "https://bertonline.sharepoint.com/sites/TestPnPSC_12345_d644f1c6-80ac-4858-8e63-a7a5ce26c206";
            //centralSubSiteUrl = "https://bertonline.sharepoint.com/sites/TestPnPSC_12345_d644f1c6-80ac-4858-8e63-a7a5ce26c206/sub";
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
        /// Site TermGroup Test
        /// </summary>
        [TestMethod]
        [Timeout(15 * 60 * 1000)]
        public void SiteCollectionTermGroupTest()
        {
            new TermGroupImplementation().SiteCollectionTermGroup(centralSiteCollectionUrl);
        }
        #endregion

        #region Web test cases
        /// <summary>
        /// Web TermGroup test
        /// </summary>
        [TestMethod]
        [Timeout(15 * 60 * 1000)]
        public void WebTermGroupTest()
        {
            new TermGroupImplementation().SiteCollectionTermGroup(centralSubSiteUrl);
        }
        #endregion
    }
}
