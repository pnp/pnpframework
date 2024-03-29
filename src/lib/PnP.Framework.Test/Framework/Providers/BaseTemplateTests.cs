﻿
using Microsoft.Online.SharePoint.TenantAdministration;
using Microsoft.SharePoint.Client;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using PnP.Framework.Entities;
using PnP.Framework.Graph;
using PnP.Framework.Provisioning.Connectors;
using PnP.Framework.Provisioning.Model;
using PnP.Framework.Provisioning.ObjectHandlers;
using PnP.Framework.Provisioning.Providers.Xml;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Xml;
using System.Xml.XPath;

namespace PnP.Framework.Test.Framework.Providers
{
    [TestClass]
    public class BaseTemplateTests
    {
        protected class BaseTemplate
        {
            public BaseTemplate(string template, string subSiteTemplate = "", string saveAsTemplate = "", bool skipDeleteCreateCycle = false)
            {
                Template = template;
                SubSiteTemplate = subSiteTemplate;
                SaveAsTemplate = saveAsTemplate;
                SkipDeleteCreateCycle = skipDeleteCreateCycle;
            }

            public string Template { get; set; }
            public string SubSiteTemplate { get; set; }
            public string SaveAsTemplate { get; set; }
            public bool SkipDeleteCreateCycle { get; set; }
        }

        [TestMethod]
        [Ignore]
        public void ExtractMinimalTemplateSetForRelease()
        {
            bool deleteSites = true;
            bool createSites = true;

            List<BaseTemplate> templates = new List<BaseTemplate>(1);
            templates.Add(new BaseTemplate("STS#0"));
            templates.Add(new BaseTemplate("STS#3"));
            templates.Add(new BaseTemplate("SITEPAGEPUBLISHING#0"));
            templates.Add(new BaseTemplate("GROUP#0", skipDeleteCreateCycle: true));
            ProcessBaseTemplates(templates, deleteSites, createSites);
        }

        [TestMethod]
        [Ignore]
        public void ExtractOtherTemplatesForRelease()
        {
            // IMPORTANT: extraction needs to be done with user credentials, not app-only. With app-only certain templates (like BDR) will fail
            // use these flags to save time if the process failed after delete or create sites was done
            bool deleteSites = false;
            bool createSites = false;

            List<BaseTemplate> templates = new List<BaseTemplate>(15);
            templates.Add(new BaseTemplate("BLOG#0"));
            templates.Add(new BaseTemplate("BDR#0"));
            templates.Add(new BaseTemplate("DEV#0"));
            templates.Add(new BaseTemplate("OFFILE#1"));
            templates.Add(new BaseTemplate("EHS#1"));
            templates.Add(new BaseTemplate("BLANKINTERNETCONTAINER#0", "", "BLANKINTERNET#0"));
            templates.Add(new BaseTemplate("BICENTERSITE#0"));
            templates.Add(new BaseTemplate("SRCHCEN#0"));
            templates.Add(new BaseTemplate("BLANKINTERNETCONTAINER#0", "CMSPUBLISHING#0", "CMSPUBLISHING#0"));
            templates.Add(new BaseTemplate("ENTERWIKI#0"));
            templates.Add(new BaseTemplate("PROJECTSITE#0"));
            templates.Add(new BaseTemplate("COMMUNITY#0"));
            templates.Add(new BaseTemplate("COMMUNITYPORTAL#0"));
            templates.Add(new BaseTemplate("SRCHCENTERLITE#0"));
            templates.Add(new BaseTemplate("VISPRUS#0"));

            ProcessBaseTemplates(templates, deleteSites, createSites);
        }


        [TestMethod]
        [Ignore]
        public void ExtractSingleTemplate2()
        {
            bool deleteSites = true;
            bool createSites = true;

            List<BaseTemplate> templates = new List<BaseTemplate>(1);
            templates.Add(new BaseTemplate("SITEPAGEPUBLISHING#0", skipDeleteCreateCycle: false));

            ProcessBaseTemplates(templates, deleteSites, createSites);
        }

        [TestMethod]
        [Ignore]
        public void ExtractBaseTemplates2()
        {
            // IMPORTANT: extraction needs to be done with user credentials, not app-only. With app-only certain templates (like BDR) will fail
            // use these flags to save time if the process failed after delete or create sites was done
            bool deleteSites = true;
            bool createSites = true;

            List<BaseTemplate> templates = new List<BaseTemplate>(15);
            templates.Add(new BaseTemplate("STS#0"));
            templates.Add(new BaseTemplate("BLOG#0"));
            templates.Add(new BaseTemplate("BDR#0"));
            templates.Add(new BaseTemplate("DEV#0"));
            templates.Add(new BaseTemplate("OFFILE#1"));
            templates.Add(new BaseTemplate("GROUP#0", skipDeleteCreateCycle: true));
            templates.Add(new BaseTemplate("SITEPAGEPUBLISHING#0"));
            templates.Add(new BaseTemplate("EHS#1"));
            templates.Add(new BaseTemplate("BLANKINTERNETCONTAINER#0", "", "BLANKINTERNET#0"));
            templates.Add(new BaseTemplate("STS#3"));
            templates.Add(new BaseTemplate("BICENTERSITE#0"));
            templates.Add(new BaseTemplate("SRCHCEN#0"));
            templates.Add(new BaseTemplate("BLANKINTERNETCONTAINER#0", "CMSPUBLISHING#0", "CMSPUBLISHING#0"));
            templates.Add(new BaseTemplate("ENTERWIKI#0"));
            templates.Add(new BaseTemplate("PROJECTSITE#0"));
            templates.Add(new BaseTemplate("COMMUNITY#0"));
            templates.Add(new BaseTemplate("COMMUNITYPORTAL#0"));
            templates.Add(new BaseTemplate("SRCHCENTERLITE#0"));
            templates.Add(new BaseTemplate("VISPRUS#0"));

            ProcessBaseTemplates(templates, deleteSites, createSites);
        }


        [TestMethod]
        [Ignore]
        public void ExtractSingleBaseTemplate2()
        {
            // use these flags to save time if the process failed after delete or create sites was done
            bool deleteSites = true;
            bool createSites = true;

            List<BaseTemplate> templates = new List<BaseTemplate>(1);
            templates.Add(new BaseTemplate("GROUP#0", skipDeleteCreateCycle: true));

            ProcessBaseTemplates(templates, deleteSites, createSites);
        }

        private void ProcessBaseTemplates(List<BaseTemplate> templates, bool deleteSites, bool createSites)
        {
            using (var tenantCtx = TestCommon.CreateTenantClientContext())
            {
                tenantCtx.RequestTimeout = 1000 * 60 * 15;
                Tenant tenant = new Tenant(tenantCtx);

                if (deleteSites)
                {
                    // First delete all template site collections when in SPO
                    foreach (var template in templates)
                    {
                        string siteUrl = GetSiteUrl(template);

                        try
                        {
                            Console.WriteLine("Deleting existing site {0}", siteUrl);
                            if (template.SkipDeleteCreateCycle)
                            {
                                // Do nothing for the time being since we don't allow group deletion using app-only context
                            }
                            else
                            {
                                tenant.DeleteSiteCollection(siteUrl, false);
                            }
                        }
                        catch{ }
                    }
                }

                if (createSites)
                {
                    // Create site collections
                    foreach (var template in templates)
                    {
                        string siteUrl = GetSiteUrl(template);

                        Console.WriteLine("Creating site {0}", siteUrl);

                        if (template.SkipDeleteCreateCycle)
                        {
                            // Do nothing for the time being since we don't allow group creation using app-only context
                        }
                        else
                        {
                            bool siteExists = false;
                            if (template.SubSiteTemplate.Length > 0)
                            {
                                siteExists = tenant.SiteExists(siteUrl);
                            }

                            if (!siteExists)
                            {
                                if (template.Template.StartsWith("SITEPAGEPUBLISHING"))
                                {
                                    using (var clientContext = TestCommon.CreateClientContext())
                                    {
                                        var newCommunicationSiteCreationInfo = new PnP.Framework.Sites.CommunicationSiteCollectionCreationInformation()
                                        {
                                            Url = siteUrl,
                                            SiteDesign = PnP.Framework.Sites.CommunicationSiteDesign.Blank,
                                            Title = "Template Site",
                                            Lcid = 1033
                                        };

                                        clientContext.CreateSiteAsync(newCommunicationSiteCreationInfo).Wait();
                                    }
                                }
                                else
                                {

                                    var newSiteEntity = new Entities.SiteEntity()
                                    {
                                        Lcid = 1033,
                                        TimeZoneId = 4,
                                        SiteOwnerLogin = TestCommon.UserName,
                                        Title = "Template Site",
                                        Template = template.Template,
                                        Url = siteUrl,
                                    };

                                    tenant.CreateSiteCollection(newSiteEntity, true, true);
                                }
                            }

                            if (template.SubSiteTemplate.Length > 0)
                            {
                                using (ClientContext ctx = TestCommon.CreateClientContext())
                                {
                                    using (var sitecolCtx = ctx.Clone(siteUrl))
                                    {
                                        sitecolCtx.Web.Webs.Add(new WebCreationInformation()
                                        {
                                            Title = string.Format("template{0}", template.SubSiteTemplate),
                                            Language = 1033,
                                            Url = string.Format("template{0}", template.SubSiteTemplate.Replace("#", "")),
                                            UseSamePermissionsAsParentSite = true
                                        });
                                        sitecolCtx.ExecuteQueryRetry();
                                    }
                                }
                            }
                        }
                    }
                }
            }

            // Export the base templates
            using (ClientContext ctx = TestCommon.CreateClientContext())
            {
                foreach (var template in templates)
                {
                    string siteUrl = GetSiteUrl(template, false);

                    // Export the base templates
                    using (ClientContext cc = ctx.Clone(siteUrl))
                    {
                        cc.RequestTimeout = 1000 * 60 * 15;

                        // Specify null as base template since we do want "everything" in this case
                        ProvisioningTemplateCreationInformation creationInfo = new ProvisioningTemplateCreationInformation(cc.Web);
                        creationInfo.BaseTemplate = null;
                        // Do not extract the home page for the base templates
                        creationInfo.HandlersToProcess ^= Handlers.PageContents;

                        // Override the save name. Case is online site collection provisioned using blankinternetcontainer#0 which returns
                        // blankinternet#0 as web template using CSOM/SSOM API
                        string templateName = template.Template;
                        if (template.SaveAsTemplate.Length > 0)
                        {
                            templateName = template.SaveAsTemplate;
                        }

                        ProvisioningTemplate p = cc.Web.GetProvisioningTemplate(creationInfo);
                        if (template.SubSiteTemplate.Length > 0)
                        {
                            p.Id = String.Format("{0}template", template.SubSiteTemplate.Replace("#", ""));
                        }
                        else
                        {
                            p.Id = String.Format("{0}template", templateName.Replace("#", ""));
                        }

                        // Cleanup before saving
                        p.Security.AdditionalAdministrators.Clear();

                        // persist the template using the XML provider
                        XMLFileSystemTemplateProvider provider = new XMLFileSystemTemplateProvider(".", "");
                        if (template.SubSiteTemplate.Length > 0)
                        {
                            provider.SaveAs(p, String.Format("{0}Template.xml", template.SubSiteTemplate.Replace("#", "")));
                        }
                        else
                        {
                            provider.SaveAs(p, String.Format("{0}Template.xml", templateName.Replace("#", "")));
                        }
                    }
                }
            }
        }

        private static string GetSiteUrl(BaseTemplate template, bool siteCollectionUrl = true)
        {
            Uri devSiteUrl = new Uri(TestCommon.AppSetting("SPODevSiteUrl"));
            string baseUrl = String.Format("{0}://{1}", devSiteUrl.Scheme, devSiteUrl.DnsSafeHost);

            string siteUrl = "";
            if (template.SubSiteTemplate.Length > 0)
            {
                if (siteCollectionUrl)
                {
                    siteUrl = string.Format("{1}/sites/template{2}", template.Template.Replace("#", ""), baseUrl, template.SubSiteTemplate.Replace("#", ""));
                }
                else
                {
                    siteUrl = string.Format("{1}/sites/template{2}/template{2}", template.Template.Replace("#", ""), baseUrl, template.SubSiteTemplate.Replace("#", ""));
                }
            }
            else
            {
                siteUrl = string.Format("{1}/sites/template{0}", template.Template.Replace("#", ""), baseUrl);
            }

            return siteUrl;
        }


        /// <summary>
        /// This is not a test, merely used to dump the needed template files
        /// </summary>
        [TestMethod]
        [Ignore]
        public void DumpBaseTemplates()
        {
            using (ClientContext ctx = TestCommon.CreateClientContext())
            {
                DumpTemplate(ctx, "STS#0");
                DumpTemplate(ctx, "BLOG#0");
                DumpTemplate(ctx, "BDR#0");
                DumpTemplate(ctx, "DEV#0");
                DumpTemplate(ctx, "OFFILE#1");

                DumpTemplate(ctx, "EHS#1");
                DumpTemplate(ctx, "BLANKINTERNETCONTAINER#0", "", "BLANKINTERNET#0");
                DumpTemplate(ctx, "BICENTERSITE#0");
                DumpTemplate(ctx, "SRCHCEN#0");
                DumpTemplate(ctx, "BLANKINTERNETCONTAINER#0", "CMSPUBLISHING#0");
                DumpTemplate(ctx, "ENTERWIKI#0");
                DumpTemplate(ctx, "PROJECTSITE#0");
                DumpTemplate(ctx, "COMMUNITY#0");
                DumpTemplate(ctx, "COMMUNITYPORTAL#0");
                DumpTemplate(ctx, "SRCHCENTERLITE#0");
                DumpTemplate(ctx, "VISPRUS#0");
            }
        }

        [TestMethod]
        [Ignore]
        public void DumpSingleTemplate()
        {
            using (ClientContext ctx = TestCommon.CreateClientContext())
            {
                DumpTemplate(ctx, "STS#0");
            }
        }

        private void DumpTemplate(ClientContext ctx, string template, string subSiteTemplate = "", string saveAsTemplate = "")
        {

            Uri devSiteUrl = new Uri(TestCommon.AppSetting("SPODevSiteUrl"));
            string baseUrl = String.Format("{0}://{1}", devSiteUrl.Scheme, devSiteUrl.DnsSafeHost);

            string siteUrl = "";
            if (subSiteTemplate.Length > 0)
            {
                siteUrl = string.Format("{1}/sites/template{0}/template{2}", template.Replace("#", ""), baseUrl, subSiteTemplate.Replace("#", ""));
                var siteCollectionUrl = string.Format("{1}/sites/template{0}", template.Replace("#", ""), baseUrl);
                CreateSiteCollection(template, siteCollectionUrl);
                using (var sitecolCtx = ctx.Clone(siteCollectionUrl))
                {
                    sitecolCtx.Web.Webs.Add(new WebCreationInformation()
                    {
                        Title = string.Format("template{0}", subSiteTemplate),
                        Language = 1033,
                        Url = string.Format("template{0}", subSiteTemplate.Replace("#", "")),
                        UseSamePermissionsAsParentSite = true
                    });
                    sitecolCtx.ExecuteQueryRetry();
                }
            }
            else
            {
                siteUrl = string.Format("{1}/sites/template{0}", template.Replace("#", ""), baseUrl);
                CreateSiteCollection(template, siteUrl);
            }

            using (ClientContext cc = ctx.Clone(siteUrl))
            {
                // Specify null as base template since we do want "everything" in this case
                ProvisioningTemplateCreationInformation creationInfo = new ProvisioningTemplateCreationInformation(cc.Web);
                creationInfo.BaseTemplate = null;
                // Do not extract the home page for the base templates
                creationInfo.HandlersToProcess ^= Handlers.PageContents;

                // Override the save name. Case is online site collection provisioned using blankinternetcontainer#0 which returns
                // blankinternet#0 as web template using CSOM/SSOM API
                if (saveAsTemplate.Length > 0)
                {
                    template = saveAsTemplate;
                }

                ProvisioningTemplate p = cc.Web.GetProvisioningTemplate(creationInfo);
                if (subSiteTemplate.Length > 0)
                {
                    p.Id = String.Format("{0}template", subSiteTemplate.Replace("#", ""));
                }
                else
                {
                    p.Id = String.Format("{0}template", template.Replace("#", ""));
                }

                // Cleanup before saving
                p.Security.AdditionalAdministrators.Clear();


                XMLFileSystemTemplateProvider provider = new XMLFileSystemTemplateProvider(".", "");
                if (subSiteTemplate.Length > 0)
                {
                    provider.SaveAs(p, String.Format("{0}Template.xml", subSiteTemplate.Replace("#", "")));
                }
                else
                {
                    provider.SaveAs(p, String.Format("{0}Template.xml", template.Replace("#", "")));
                }

                using (var tenantCtx = TestCommon.CreateTenantClientContext())
                {
                    Tenant tenant = new Tenant(tenantCtx);
                    Console.WriteLine("Deleting new site {0}", string.Format("{1}/sites/template{0}", template.Replace("#", ""), baseUrl));
                    tenant.DeleteSiteCollection(siteUrl, false);
                }
            }
        }

        private static void CreateSiteCollection(string template, string siteUrl)
        {
            // check if site exists
            using (var tenantCtx = TestCommon.CreateTenantClientContext())
            {
                Tenant tenant = new Tenant(tenantCtx);

                if (tenant.SiteExists(siteUrl))
                {
                    Console.WriteLine("Deleting existing site {0}", siteUrl);
                    tenant.DeleteSiteCollection(siteUrl, false);
                }
                Console.WriteLine("Creating new site {0}", siteUrl);
                tenant.CreateSiteCollection(new Entities.SiteEntity()
                {
                    Lcid = 1033,
                    TimeZoneId = 4,
                    SiteOwnerLogin = TestCommon.UserName,
                    Title = "Template Site",
                    Template = template,
                    Url = siteUrl,
                }, true, true);

            }
        }

        /// <summary>
        /// Get the base template for the current site
        /// </summary>
        [TestMethod]
        public void GetBaseTemplateForCurrentSiteTest()
        {
            using (ClientContext ctx = TestCommon.CreateClientContext())
            {
                ProvisioningTemplate t = ctx.Web.GetBaseTemplate();

                Assert.IsNotNull(t);
            }
        }

    }
}
