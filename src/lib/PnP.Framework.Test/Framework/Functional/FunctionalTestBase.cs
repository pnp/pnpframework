using Microsoft.Online.SharePoint.TenantAdministration;
using Microsoft.SharePoint.Client;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using PnP.Framework.Entities;
using System;
using System.Collections.Generic;

namespace PnP.Framework.Test.Framework.Functional
{
    [TestClass()]
    public abstract class FunctionalTestBase
    {
        private static readonly string sitecollectionNamePrefix = "TestPnPSC_12345_";

        internal static string centralSiteCollectionUrl = "";
        internal static string centralSubSiteUrl = "";
        internal const string centralSubSiteName = "sub";
        internal static bool debugMode = false;
        internal string sitecollectionName = "";

        #region Test preparation
        public static void ClassInitBase(TestContext context, bool noScriptSite = false, bool useSTS = false)
        {
            // Drop all previously created site collections to keep the environment clean
            using (var tenantContext = TestCommon.CreateTenantClientContext())
            {
                if (!debugMode)
                {
                    CleanupAllTestSiteCollections(tenantContext);

                    // Each class inheriting from this base class gets a central test site collection, so let's create that one
                    var tenant = new Tenant(tenantContext);
                    centralSiteCollectionUrl = CreateTestSiteCollection(tenant, sitecollectionNamePrefix + Guid.NewGuid().ToString(), useSTS);

                    // Add a default sub site
                    centralSubSiteUrl = CreateTestSubSite(tenant, centralSiteCollectionUrl, centralSubSiteName);

                    // Apply noscript setting
                    if (noScriptSite)
                    {
                        Console.WriteLine("Setting site {0} as NoScript", centralSiteCollectionUrl);
                        tenant.SetSiteProperties(centralSiteCollectionUrl, noScriptSite: true);
                    }
                }
            }
        }

        public static void ClassCleanupBase()
        {
            if (!debugMode)
            {
                using (var tenantContext = TestCommon.CreateTenantClientContext())
                {
                    CleanupAllTestSiteCollections(tenantContext);
                }
            }
        }

        [TestInitialize()]
        public virtual void Initialize()
        {
            TestCommon.FixAssemblyResolving("Newtonsoft.Json");
            sitecollectionName = sitecollectionNamePrefix + Guid.NewGuid().ToString();
        }

        #endregion

        #region Helper methods
        internal static string CreateTestSiteCollection(Tenant tenant, string sitecollectionName, bool useSts)
        {
            try
            {
                string devSiteUrl = TestCommon.AppSetting("SPODevSiteUrl");
                string siteToCreateUrl = GetTestSiteCollectionName(devSiteUrl, sitecollectionName);

                string siteOwnerLogin = TestCommon.AppSetting("SPOUserName");
                if (TestCommon.AppOnlyTesting())
                {
                    using (var clientContext = TestCommon.CreateClientContext())
                    {
                        List<UserEntity> admins = clientContext.Web.GetAdministrators();
                        siteOwnerLogin = admins[0].LoginName.Split(new string[] { "|" }, StringSplitOptions.RemoveEmptyEntries)[2];
                    }
                }

                if (useSts)
                {
                    SiteEntity siteToCreate = new SiteEntity()
                    {
                        Url = siteToCreateUrl,
                        Template = "STS#0",
                        Title = "Test",
                        Description = "Test site collection",
                        SiteOwnerLogin = siteOwnerLogin,
                        Lcid = 1033,
                        StorageMaximumLevel = 100,
                        UserCodeMaximumLevel = 0
                    };

                    tenant.CreateSiteCollection(siteToCreate, false, true);
                }
                else
                {
                    var commResults = (tenant.Context.Clone(devSiteUrl) as ClientContext).CreateSiteAsync(new PnP.Framework.Sites.CommunicationSiteCollectionCreationInformation()
                    {
                        Url = siteToCreateUrl,
                        SiteDesign = PnP.Framework.Sites.CommunicationSiteDesign.Blank,
                        Title = "Test",
                        Owner = siteOwnerLogin,
                        Lcid = 1033
                    }).GetAwaiter().GetResult();
                }

                return siteToCreateUrl;
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToDetailedString(tenant.Context));
                throw;
            }
        }

        private static void CleanupAllTestSiteCollections(ClientContext tenantContext)
        {
            var tenant = new Tenant(tenantContext);

            try
            {
                var siteCols = tenant.GetSiteCollections();

                foreach (var siteCol in siteCols)
                {
                    if (siteCol.Url.Contains(sitecollectionNamePrefix))
                    {
                        try
                        {
                            // Drop the site collection from the recycle bin
                            if (tenant.CheckIfSiteExists(siteCol.Url, "Recycled"))
                            {
                                tenant.DeleteSiteCollectionFromRecycleBin(siteCol.Url, false);
                            }
                            else
                            {
                                // Eat the exceptions: would occur if the site collection is already in the recycle bin.
                                try
                                {
                                    // ensure the site collection in unlocked state before deleting
                                    tenant.SetSiteLockState(siteCol.Url, SiteLockState.Unlock);
                                }
                                catch { }

                                // delete the site collection, do not use the recyle bin
                                tenant.DeleteSiteCollection(siteCol.Url, false);
                            }
                        }
                        catch (Exception ex)
                        {
                            // eat all exceptions
                            Console.WriteLine(ex.ToDetailedString(tenant.Context));
                        }
                    }
                }
            }
            // catch exceptions with the GetSiteCollections call and log them so we can grab the corelation ID
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToDetailedString(tenant.Context));
                throw;
            }

        }

        internal static string CreateTestSubSite(Tenant tenant, string sitecollectionUrl, string subSiteName)
        {
            try
            {
                // Create a sub site in the central site collection
                using (var cc = TestCommon.CreateClientContext(sitecollectionUrl))
                {
                    //Create sub site
                    SiteEntity sub = new SiteEntity() { Title = "Sub site for engine testing", Url = subSiteName, Description = "" };
                    var subWeb = cc.Web.CreateWeb(sub);
                    subWeb.EnsureProperty(t => t.Url);
                    return subWeb.Url;
                }

            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToDetailedString());
                throw;
            }

            // Below approach is not working on edog...to be investigated
            //// create a sub site in the central site collection
            //Site site = tenant.GetSiteByUrl(sitecollectionUrl);
            //tenant.Context.Load(site);
            //tenant.Context.ExecuteQueryRetry();
            //Web web = site.RootWeb;
            //web.Context.Load(web);
            //web.Context.ExecuteQueryRetry();

            ////Create sub site
            //SiteEntity sub = new SiteEntity() { Title = "Sub site for engine testing", Url = subSiteName, Description = "" };
            //var subWeb = web.CreateWeb(sub);
            //subWeb.EnsureProperty(t => t.Url);
            //return subWeb.Url;
        }

        private static string GetTestSiteCollectionName(string devSiteUrl, string siteCollection)
        {
            Uri u = new Uri(devSiteUrl);
            string host = String.Format("{0}://{1}", u.Scheme, u.DnsSafeHost);

            string path = u.AbsolutePath;
            if (path.EndsWith("/"))
            {
                path = path.Substring(0, path.Length - 1);
            }
            path = path.Substring(0, path.LastIndexOf('/'));

            return string.Format("{0}{1}/{2}", host, path, siteCollection);
        }
        #endregion

    }
}
