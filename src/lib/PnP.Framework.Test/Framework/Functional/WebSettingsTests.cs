using Microsoft.SharePoint.Client;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using PnP.Framework.Enums;
using PnP.Framework.Provisioning.Model;
using PnP.Framework.Provisioning.ObjectHandlers;
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
    /// Test cases for the provisioning engine web settings functionality
    /// </summary>
    [TestClass]
    public class WebSettingsTests : FunctionalTestBase
    {
        #region Construction
        public WebSettingsTests()
        {
            //debugMode = true;
            //centralSiteCollectionUrl = "https://bertonline.sharepoint.com/sites/TestPnPSC_12345_f449c481-ce49-4185-9ba1-f30c1752552c";
            //centralSubSiteUrl = "https://bertonline.sharepoint.com/sites/TestPnPSC_12345_f449c481-ce49-4185-9ba1-f30c1752552c/sub";
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
        /// Site WebSettings Test
        /// </summary>
        [TestMethod]
        [Timeout(15 * 60 * 1000)]
        public void SiteCollectionWebSettingsTest()
        {
            new WebSettingsImplementation().SiteCollectionWebSettings(centralSiteCollectionUrl);
        }

        /// <summary>
        /// Site Auditsettings Test
        /// </summary>
        [TestMethod]
        [Timeout(15 * 60 * 1000)]
        public void SiteCollectionAuditSettingsTest()
        {
            new WebSettingsImplementation().SiteCollectionAuditSettings(centralSiteCollectionUrl);
        }
        #endregion

        #region Web test cases
        /// <summary>
        /// Web WebSettings test
        /// </summary>
        [TestMethod]
        [Timeout(15 * 60 * 1000)]
        public void WebWebSettingsTest()
        {
            new WebSettingsImplementation().WebWebSettings(centralSiteCollectionUrl, centralSubSiteUrl);
        }

        // Audit settings are only possible on site collection level, hence no test at web level!
        #endregion
    }
}
