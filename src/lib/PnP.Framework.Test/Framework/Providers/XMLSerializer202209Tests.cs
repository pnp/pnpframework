using Microsoft.SharePoint.Client;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using PnP.Framework.Provisioning.Model;
using PnP.Framework.Provisioning.Model.AzureActiveDirectory;
using PnP.Framework.Provisioning.Model.Teams;
using PnP.Framework.Provisioning.Providers.Xml;
using PnP.Framework.Utilities;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Security;
using System.Xml.Linq;
using App = PnP.Framework.Provisioning.Model.App;
using CalendarType = Microsoft.SharePoint.Client.CalendarType;
using CanvasSectionType = PnP.Framework.Provisioning.Model.CanvasSectionType;
using ClientSidePageHeaderLayoutType = PnP.Framework.Provisioning.Model.ClientSidePageHeaderLayoutType;
using ClientSidePageHeaderTextAlignment = PnP.Framework.Provisioning.Model.ClientSidePageHeaderTextAlignment;
using ClientSidePageHeaderType = PnP.Framework.Provisioning.Model.ClientSidePageHeaderType;
using DayOfWeek = System.DayOfWeek;
using DocumentSetTemplate = PnP.Framework.Provisioning.Model.DocumentSetTemplate;
using File = System.IO.File;
using FileLevel = PnP.Framework.Provisioning.Model.FileLevel;
using ProvisioningTemplate = PnP.Framework.Provisioning.Model.ProvisioningTemplate;
using TeamTemplate = PnP.Framework.Provisioning.Model.Teams.TeamTemplate;
using WorkHour = PnP.Framework.Provisioning.Model.WorkHour;

// Replace the following statements to update the test to a new schema version
using PnP.Framework.Provisioning.Providers.Xml.V202209;
using TargetNamespace = PnP.Framework.Provisioning.Providers.Xml.V202209;
using TargetSerializer = PnP.Framework.Provisioning.Providers.Xml.XMLPnPSchemaV202209Serializer;
using TargetProvisioning = PnP.Framework.Provisioning.Providers.Xml.V202209.Provisioning;

namespace PnP.Framework.Test.Framework.Providers
{
    /// <summary>
    /// Covers below objects:
    /// ProvisioningTemplate:
    ///     Properties
    ///     SitePolicy
    ///     WebSettings 
    ///     RegionalSettings
    ///     SupportedUILanguages
    ///     AuditSettings
    ///     PropertyBagEntries
    ///     Security
    ///     Navigation
    ///     SiteFields
    ///     ContentTypes
    ///     Lists
    ///     Features
    ///     CustomActions
    ///     Files
    ///     Pages
    ///     TermGroups
    ///     ComposedLook
    ///     SearchSettings
    ///     Publishing
    ///     SiteWebhooks
    ///     ClientSidePages
    ///     ALM
    ///     Header
    ///     Footer
    ///     ProvisioningTemplateWebhooks 
    ///     SiteSettings
    /// Teams:
    ///     TeamTemplate
    ///     Team
    ///     Apps
    /// AzureActiveDirectory:
    ///     Users
    /// Tenant:
    ///     AppCatalog
    ///     WebApiPermissions
    ///     ContentDeliveryNetwork
    ///     SiteDesigns
    ///     SiteScripts
    ///     StorageEntities
    ///     Themes
    ///     SPUserProfile
    ///     Office365Group Settings
    ///     Office365Group Lifecycle
    ///     Sharing Settings
    /// Drive
    /// Localizations
    /// 
    /// </summary>
    [TestClass]
    public class XMLSerializer202209Tests
    {
        #region Test variables
        private const string TEST_CATEGORY = "Framework Provisioning XML Serialization\\Deserialization 202209";
        private const string TEST_OUT_FILE = "ProvisioningTemplate-2022-09-Sample-01-test.xml";
        private const string TEST_TEMPLATE = "ProvisioningSchema-2022-09-FullSample-01.xml";

        #endregion

        #region Test initialize
        [ClassCleanup]
        public static void Clean()
        {
            var testFilePath = Path.GetFullPath(Path.Combine("../../Resources/Templates", TEST_OUT_FILE));
            if (File.Exists(testFilePath))
            {
                File.Delete(testFilePath);
            }
        }
        #endregion

        [TestMethod]
        [TestCategory(TEST_CATEGORY)]
        public void XMLSerializer_Deserialize_Apps()
        {
            var provider = new XMLFileSystemTemplateProvider($@"{AppDomain.CurrentDomain.BaseDirectory}\..\..\..\Resources", "Templates");

            var serializer = new TargetSerializer();
            var template = provider.GetTemplate(TEST_TEMPLATE, serializer);

            Assert.AreEqual(3, template.ApplicationLifecycleManagement.Apps.Count);
            Assert.AreEqual(AppAction.InstallOrUpdate, template.ApplicationLifecycleManagement.Apps.First().Action);
            Assert.AreEqual("d0816f0a-fda4-4a98-8e61-1bbe1f2b5b27", template.ApplicationLifecycleManagement.Apps.First().AppId);
            Assert.AreEqual(SyncMode.Synchronously, template.ApplicationLifecycleManagement.Apps.First().SyncMode);
            Assert.AreEqual(AppAction.Update, template.ApplicationLifecycleManagement.Apps[1].Action);
            Assert.AreEqual(AppAction.Uninstall, template.ApplicationLifecycleManagement.Apps[2].Action);
        }

        [TestMethod]
        [TestCategory(TEST_CATEGORY)]
        public void XMLSerializer_Serialize_Apps()
        {
            var provider = new XMLFileSystemTemplateProvider($@"{AppDomain.CurrentDomain.BaseDirectory}\..\..\..\Resources", "Templates");

            var result = new ProvisioningTemplate();

            result.ApplicationLifecycleManagement.Apps.Add(new App
            {
                AppId = "1d602ad8-2ef2-4d0b-bfdc-ca71a1d591fc",
                Action = AppAction.Install,
                SyncMode = SyncMode.Asynchronously
            });

            result.ApplicationLifecycleManagement.Apps.Add(new App
            {
                AppId = "6f9a1b33-a13a-4313-b106-0effa15624e6",
                Action = AppAction.Uninstall,
                SyncMode = SyncMode.Synchronously
            });

            result.ApplicationLifecycleManagement.Apps.Add(new App
            {
                AppId = "a44cd745-57ea-44bc-a707-95b985e696e5",
                Action = AppAction.Update,
                SyncMode = SyncMode.Synchronously
            });

            var serializer = new TargetSerializer();
            provider.SaveAs(result, TEST_OUT_FILE, serializer);

            var path = $"{provider.Connector.Parameters["ConnectionString"]}\\{provider.Connector.Parameters["Container"]}\\{TEST_OUT_FILE}";
            Assert.IsTrue(File.Exists(path)); path = Path.GetFullPath(path);
            var xml = XDocument.Load(path);
            TargetProvisioning wrappedResult =
                XMLSerializer.Deserialize<TargetProvisioning>(xml);

            var template = wrappedResult.Templates[0].ProvisioningTemplate.First();
            Assert.AreEqual(3, template.ApplicationLifecycleManagement.Apps.Count());
            Assert.AreEqual("1d602ad8-2ef2-4d0b-bfdc-ca71a1d591fc", template.ApplicationLifecycleManagement.Apps[0].AppId);
            Assert.AreEqual(ApplicationLifecycleManagementAppAction.Install, template.ApplicationLifecycleManagement.Apps[0].Action);
            Assert.AreEqual(ApplicationLifecycleManagementAppAction.Uninstall, template.ApplicationLifecycleManagement.Apps[1].Action);
            Assert.AreEqual(ApplicationLifecycleManagementAppAction.Update, template.ApplicationLifecycleManagement.Apps[2].Action);
            Assert.AreEqual(ApplicationLifecycleManagementAppSyncMode.Asynchronously, template.ApplicationLifecycleManagement.Apps[0].SyncMode);
            Assert.AreEqual(ApplicationLifecycleManagementAppSyncMode.Synchronously, template.ApplicationLifecycleManagement.Apps[1].SyncMode);
        }

        [TestMethod]
        [TestCategory(TEST_CATEGORY)]
        public void XMLSerializer_Deserialize_SiteHeader()
        {
            var provider = new XMLFileSystemTemplateProvider($@"{AppDomain.CurrentDomain.BaseDirectory}\..\..\..\Resources", "Templates");

            var serializer = new TargetSerializer();
            var template = provider.GetTemplate(TEST_TEMPLATE, serializer);

            Assert.AreEqual(SiteHeaderLayout.Standard, template.Header.Layout);
            Assert.AreEqual(SiteHeaderMenuStyle.MegaMenu, template.Header.MenuStyle);
            Assert.IsTrue(template.Header.ShowSiteTitle);
            Assert.IsTrue(template.Header.ShowSiteNavigation);
            Assert.AreEqual(PnP.Framework.Provisioning.Model.Emphasis.Soft, template.Header.BackgroundEmphasis);
        }

        [TestMethod]
        [TestCategory(TEST_CATEGORY)]
        public void XMLSerializer_Serialize_SiteHeader()
        {
            var provider = new XMLFileSystemTemplateProvider($@"{AppDomain.CurrentDomain.BaseDirectory}\..\..\..\Resources", "Templates");

            var result = new ProvisioningTemplate
            {
                Header = new SiteHeader
                {
                    MenuStyle = SiteHeaderMenuStyle.Cascading,
                    Layout = SiteHeaderLayout.Compact,
                    BackgroundEmphasis = PnP.Framework.Provisioning.Model.Emphasis.Strong,
                    ShowSiteTitle = true,
                    ShowSiteNavigation = true,
                }
            };

            var serializer = new TargetSerializer();
            provider.SaveAs(result, TEST_OUT_FILE, serializer);

            var path = $"{provider.Connector.Parameters["ConnectionString"]}\\{provider.Connector.Parameters["Container"]}\\{TEST_OUT_FILE}";
            Assert.IsTrue(File.Exists(path)); path = Path.GetFullPath(path);
            var xml = XDocument.Load(path);
            var wrappedResult =
                XMLSerializer.Deserialize<TargetProvisioning>(xml);

            var template = wrappedResult.Templates[0].ProvisioningTemplate.First();

            Assert.AreEqual(HeaderLayout.Compact, template.Header.Layout);
            Assert.AreEqual(HeaderMenuStyle.Cascading, template.Header.MenuStyle);
            Assert.IsTrue(template.Header.ShowSiteTitle);
            Assert.IsTrue(template.Header.ShowSiteNavigation);
            Assert.AreEqual(TargetNamespace.Emphasis.Strong, template.Header.BackgroundEmphasis);
        }

        [TestMethod]
        [TestCategory(TEST_CATEGORY)]
        public void XMLSerializer_Deserialize_TemplateTheme()
        {
            var provider = new XMLFileSystemTemplateProvider($@"{AppDomain.CurrentDomain.BaseDirectory}\..\..\..\Resources", "Templates");

            var serializer = new TargetSerializer();
            var template = provider.GetTemplate(TEST_TEMPLATE, serializer);

            Assert.AreEqual(false, template.Theme.IsInverted);
            Assert.AreEqual(false, template.Theme.Overwrite);
            Assert.AreEqual("CustomOrange", template.Theme.Name);
            Assert.IsTrue(template.Theme.Palette.Contains("\"neutralQuaternaryAlt\": \"#dadada\""));
        }

        [TestMethod]
        [TestCategory(TEST_CATEGORY)]
        public void XMLSerializer_Serialize_TemplateTheme()
        {
            var provider = new XMLFileSystemTemplateProvider($@"{AppDomain.CurrentDomain.BaseDirectory}\..\..\..\Resources", "Templates");

            var result = new ProvisioningTemplate
            {
                Theme = new PnP.Framework.Provisioning.Model.Theme
                {
                    Name = "CustomOrange",
                    IsInverted = false,
                    Palette = "{\"neutralQuaternaryAlt\": \"#dadada\"}",
                    Overwrite = false
                }
            };

            var serializer = new TargetSerializer();
            provider.SaveAs(result, TEST_OUT_FILE, serializer);

            var path = $"{provider.Connector.Parameters["ConnectionString"]}\\{provider.Connector.Parameters["Container"]}\\{TEST_OUT_FILE}";
            Assert.IsTrue(File.Exists(path)); path = Path.GetFullPath(path);
            var xml = XDocument.Load(path);
            var wrappedResult =
                XMLSerializer.Deserialize<TargetProvisioning>(xml);

            var template = wrappedResult.Templates[0].ProvisioningTemplate.First();

            Assert.AreEqual(false, template.Theme.IsInverted);
            Assert.AreEqual(false, template.Theme.Overwrite);
            Assert.AreEqual("CustomOrange", template.Theme.Name);
            Assert.IsTrue(template.Theme.Text[0].Contains("\"neutralQuaternaryAlt\": \"#dadada\""));
        }

        [TestMethod]
        [TestCategory(TEST_CATEGORY)]
        public void XMLSerializer_Deserialize_SiteFooter()
        {
            var provider = new XMLFileSystemTemplateProvider($@"{AppDomain.CurrentDomain.BaseDirectory}\..\..\..\Resources", "Templates");

            var serializer = new TargetSerializer();
            var template = provider.GetTemplate(TEST_TEMPLATE, serializer);

            Assert.AreEqual(true, template.Footer.Enabled);
            Assert.AreEqual("logo.png", template.Footer.Logo);
            Assert.AreEqual("FooterName", template.Footer.Name);
            Assert.AreEqual(true, template.Footer.RemoveExistingNodes);
            Assert.AreEqual(SiteFooterLayout.Simple, template.Footer.Layout);
            Assert.AreEqual(Provisioning.Model.Emphasis.Soft, template.Footer.BackgroundEmphasis);
            Assert.AreEqual("Custom footer", template.Footer.DisplayName);
            Assert.AreEqual(6, template.Footer.FooterLinks.Count);
            Assert.AreEqual("www.link01.com", template.Footer.FooterLinks[0].Url);
            Assert.AreEqual("Link 01", template.Footer.FooterLinks[0].DisplayName);
            Assert.AreEqual(3, template.Footer.FooterLinks[3].FooterLinks.Count);
            Assert.AreEqual("www.link04-01.com", template.Footer.FooterLinks[3].FooterLinks[0].Url);
            Assert.AreEqual("Child Link 04-01", template.Footer.FooterLinks[3].FooterLinks[0].DisplayName);
            Assert.IsNull(template.Footer.FooterLinks[3].Url);
            Assert.AreEqual(2, template.Footer.FooterLinks[5].FooterLinks[0].FooterLinks.Count);
        }

        [TestMethod]
        [TestCategory(TEST_CATEGORY)]
        public void XMLSerializer_Serialize_SiteFooter()
        {
            var provider = new XMLFileSystemTemplateProvider($@"{AppDomain.CurrentDomain.BaseDirectory}\..\..\..\Resources", "Templates");

            var result = new ProvisioningTemplate
            {
                Footer = new SiteFooter
                {
                    Enabled = true,
                    Logo = "logo.png",
                    Name = "MyFooter",
                    RemoveExistingNodes = true,
                    DisplayName = "MyFooter",
                    Layout = SiteFooterLayout.Extended,
                    BackgroundEmphasis = Provisioning.Model.Emphasis.Soft,
                    FooterLinks = {
                        new SiteFooterLink
                        {
                            Url = "www.link01.com",
                            DisplayName = "Link 01"
                        },
                        new SiteFooterLink
                        {
                            DisplayName = "Link 02",
                            FooterLinks =
                            {
                                new SiteFooterLink
                                {
                                    Url = "www.link02-01.com",
                                    DisplayName = "Child Link 01",
                                },
                                new SiteFooterLink
                                {
                                    Url = "www.link02-02.com",
                                    DisplayName = "Child Link 02",
                                }
                            }
                        }
                    }
                }
            };

            var serializer = new TargetSerializer();
            provider.SaveAs(result, TEST_OUT_FILE, serializer);

            var path = $"{provider.Connector.Parameters["ConnectionString"]}\\{provider.Connector.Parameters["Container"]}\\{TEST_OUT_FILE}";
            Assert.IsTrue(File.Exists(path)); path = Path.GetFullPath(path);
            var xml = XDocument.Load(path);
            var wrappedResult =
                XMLSerializer.Deserialize<TargetProvisioning>(xml);

            var template = wrappedResult.Templates[0].ProvisioningTemplate.First();

            Assert.AreEqual(2, template.Footer.FooterLinks.Count());
            Assert.AreEqual(true, template.Footer.Enabled);
            Assert.AreEqual("logo.png", template.Footer.Logo);
            Assert.AreEqual("MyFooter", template.Footer.Name);
            Assert.AreEqual(true, template.Footer.RemoveExistingNodes);
            Assert.AreEqual("MyFooter", template.Footer.DisplayName);
            Assert.AreEqual(FooterLayout.Extended, template.Footer.Layout);
            Assert.AreEqual(TargetNamespace.Emphasis.Soft, template.Footer.BackgroundEmphasis);
            Assert.AreEqual("www.link01.com", template.Footer.FooterLinks[0].Url);
            Assert.AreEqual("Link 01", template.Footer.FooterLinks[0].DisplayName);
            Assert.AreEqual(2, template.Footer.FooterLinks[1].FooterLink1.Count());
            Assert.AreEqual("www.link02-01.com", template.Footer.FooterLinks[1].FooterLink1[0].Url);
            Assert.AreEqual("Child Link 01", template.Footer.FooterLinks[1].FooterLink1[0].DisplayName);
        }

        [TestMethod]
        [TestCategory(TEST_CATEGORY)]
        public void XMLSerializer_Deserialize_AppCatalog()
        {
            var provider = new XMLFileSystemTemplateProvider($@"{AppDomain.CurrentDomain.BaseDirectory}\..\..\..\Resources", "Templates");

            var serializer = new TargetSerializer();
            var template = provider.GetTemplate(TEST_TEMPLATE, serializer);

            Assert.AreEqual(3, template.ApplicationLifecycleManagement.AppCatalog.Packages.Count);
            Assert.AreEqual(true, template.ApplicationLifecycleManagement.AppCatalog.Packages[0].Overwrite);
            Assert.AreEqual(true, template.ApplicationLifecycleManagement.AppCatalog.Packages[0].SkipFeatureDeployment);
            Assert.AreEqual(PackageAction.UploadAndPublish, template.ApplicationLifecycleManagement.AppCatalog.Packages[0].Action);
            Assert.AreEqual(PackageAction.Publish, template.ApplicationLifecycleManagement.AppCatalog.Packages[1].Action);
            Assert.AreEqual("solution\\spfx-discuss-now.sppkg", template.ApplicationLifecycleManagement.AppCatalog.Packages[0].Src);
            Assert.AreEqual("9672a07b-c111-4a12-bb5b-8d43c2ddd256", template.ApplicationLifecycleManagement.AppCatalog.Packages[2].PackageId);

            Assert.AreEqual(3, template.ApplicationLifecycleManagement.Apps.Count);
            Assert.AreEqual("d0816f0a-fda4-4a98-8e61-1bbe1f2b5b27", template.ApplicationLifecycleManagement.Apps[0].AppId);
            Assert.AreEqual(AppAction.InstallOrUpdate, template.ApplicationLifecycleManagement.Apps[0].Action);
            Assert.AreEqual(SyncMode.Synchronously, template.ApplicationLifecycleManagement.Apps[0].SyncMode);

            Assert.AreEqual(3, template.Tenant.AppCatalog.Packages.Count);
            Assert.AreEqual(true, template.Tenant.AppCatalog.Packages[0].Overwrite);
            Assert.AreEqual(true, template.Tenant.AppCatalog.Packages[0].SkipFeatureDeployment);
            Assert.AreEqual(PackageAction.UploadAndPublish, template.Tenant.AppCatalog.Packages[0].Action);
            Assert.AreEqual(PackageAction.Publish, template.Tenant.AppCatalog.Packages[1].Action);
            Assert.AreEqual("solution\\spfx-discuss-now.sppkg", template.Tenant.AppCatalog.Packages[0].Src);
            Assert.AreEqual("9672a07b-c111-4a12-bb5b-8d43c2ddd256", template.Tenant.AppCatalog.Packages[2].PackageId);
        }

        [TestMethod]
        [TestCategory(TEST_CATEGORY)]
        public void XMLSerializer_Serialize_AppCatalog()
        {
            var provider = new XMLFileSystemTemplateProvider($@"{AppDomain.CurrentDomain.BaseDirectory}\..\..\..\Resources", "Templates");

            var result = new ProvisioningTemplate();

            var packages = new PackageCollection(result)
            {
                new Package
                {
                    SkipFeatureDeployment = true,
                    Src = "testpackage.sppkg",
                    Overwrite = true,
                    Action = PackageAction.Upload,
                    PackageId = "60006518-60b3-46d1-8aa7-60a89ce35f03"
                },
                new Package
                {
                    SkipFeatureDeployment = true,
                    Overwrite = true,
                    Action = PackageAction.Publish,
                    PackageId = "60006518-60b3-46d1-8aa7-60a89ce35f03"
                }
            };

            var apps = new AppCollection(result)
            {
                new App
                {
                    AppId = "d0816f0a-fda4-4a98-8e61-1bbe1f2b5b27",
                    Action = AppAction.InstallOrUpdate,
                    SyncMode = SyncMode.Synchronously
                },
                new App
                {
                    AppId = "{AppPackageId:spfx-discuss-now}",
                    Action = AppAction.Update,
                    SyncMode = SyncMode.Asynchronously
                },
                new App
                {
                    AppId = "9672a07b-c111-4a12-bb5b-8d43c2ddd256",
                    Action = AppAction.Uninstall
                },
            };

            result.ApplicationLifecycleManagement = new PnP.Framework.Provisioning.Model.ApplicationLifecycleManagement();

            result.ApplicationLifecycleManagement.AppCatalog.Packages.AddRange(packages);
            result.ApplicationLifecycleManagement.Apps.AddRange(apps);
            result.Tenant = new ProvisioningTenant(result.ApplicationLifecycleManagement.AppCatalog, new PnP.Framework.Provisioning.Model.ContentDeliveryNetwork());

            var serializer = new TargetSerializer();
            provider.SaveAs(result, TEST_OUT_FILE, serializer);

            var path = $"{provider.Connector.Parameters["ConnectionString"]}\\{provider.Connector.Parameters["Container"]}\\{TEST_OUT_FILE}";
            Assert.IsTrue(File.Exists(path)); path = Path.GetFullPath(path);
            var xml = XDocument.Load(path);
            var wrappedResult = XMLSerializer.Deserialize<TargetProvisioning>(xml);

            var template = wrappedResult.Templates[0].ProvisioningTemplate.First();

            Assert.AreEqual(2, template.ApplicationLifecycleManagement.AppCatalog.Count());
            Assert.AreEqual(true, template.ApplicationLifecycleManagement.AppCatalog[0].SkipFeatureDeployment);
            Assert.AreEqual(true, template.ApplicationLifecycleManagement.AppCatalog[0].Overwrite);
            Assert.AreEqual("testpackage.sppkg", template.ApplicationLifecycleManagement.AppCatalog[0].Src); Assert.AreEqual("60006518-60b3-46d1-8aa7-60a89ce35f03", template.ApplicationLifecycleManagement.AppCatalog[1].PackageId);
            Assert.AreEqual(AppCatalogPackageAction.Publish, template.ApplicationLifecycleManagement.AppCatalog[1].Action);

            Assert.AreEqual(3, template.ApplicationLifecycleManagement.Apps.Count());
            Assert.AreEqual("d0816f0a-fda4-4a98-8e61-1bbe1f2b5b27", template.ApplicationLifecycleManagement.Apps[0].AppId);
            Assert.AreEqual(ApplicationLifecycleManagementAppAction.InstallOrUpdate, template.ApplicationLifecycleManagement.Apps[0].Action);
            Assert.AreEqual(ApplicationLifecycleManagementAppSyncMode.Synchronously, template.ApplicationLifecycleManagement.Apps[0].SyncMode);

            Assert.AreEqual(2, wrappedResult.Tenant.AppCatalog.Count());
            Assert.AreEqual(true, wrappedResult.Tenant.AppCatalog[0].SkipFeatureDeployment);
            Assert.AreEqual(true, wrappedResult.Tenant.AppCatalog[0].Overwrite);
            Assert.AreEqual("testpackage.sppkg", wrappedResult.Tenant.AppCatalog[0].Src); Assert.AreEqual("60006518-60b3-46d1-8aa7-60a89ce35f03", wrappedResult.Tenant.AppCatalog[1].PackageId);
            Assert.AreEqual(AppCatalogPackageAction.Publish, wrappedResult.Tenant.AppCatalog[1].Action);
        }

        [TestMethod]
        [TestCategory(TEST_CATEGORY)]
        public void XMLSerializer_Deserialize_ProvisioningTemplateWebhook()
        {
            var provider = new XMLFileSystemTemplateProvider($@"{AppDomain.CurrentDomain.BaseDirectory}\..\..\..\Resources", "Templates");

            var serializer = new TargetSerializer();
            var template = provider.GetTemplate(TEST_TEMPLATE, serializer);

            Assert.AreEqual(2, template.ProvisioningTemplateWebhooks.Count);
            Assert.IsTrue(template.ProvisioningTemplateWebhooks[0].Url.Contains("https://mywebhook.azurefunctions.net/function01"));
            Assert.AreEqual(ProvisioningTemplateWebhookMethod.GET, template.ProvisioningTemplateWebhooks[0].Method);
            Assert.AreEqual(ProvisioningTemplateWebhookKind.ProvisioningTemplateStarted, template.ProvisioningTemplateWebhooks[0].Kind);
            Assert.AreEqual(ProvisioningTemplateWebhookBodyFormat.Json, template.ProvisioningTemplateWebhooks[1].BodyFormat);
            Assert.AreEqual(true, template.ProvisioningTemplateWebhooks[1].Async);
            Assert.AreEqual(3, template.ProvisioningTemplateWebhooks[0].Parameters.Count);
            Assert.IsTrue(template.ProvisioningTemplateWebhooks[0].Parameters.ContainsKey("Param01"));
            Assert.AreEqual("Value01", template.ProvisioningTemplateWebhooks[0].Parameters["Param01"]);
            Assert.AreEqual("{sitecollection}", template.ProvisioningTemplateWebhooks[1].Parameters["Site"]);
        }

        [TestMethod]
        [TestCategory(TEST_CATEGORY)]
        public void XMLSerializer_Serialize_ProvisioningTemplateWebhook()
        {
            var provider = new XMLFileSystemTemplateProvider($@"{AppDomain.CurrentDomain.BaseDirectory}\..\..\..\Resources", "Templates");

            var result = new ProvisioningTemplate();

            result.ProvisioningTemplateWebhooks.Add(new PnP.Framework.Provisioning.Model.ProvisioningTemplateWebhook
            {
                Url = "https://my.url/func01",
                Method = ProvisioningTemplateWebhookMethod.GET,
                Async = false,
                Kind = ProvisioningTemplateWebhookKind.ObjectHandlerProvisioningStarted,
                BodyFormat = ProvisioningTemplateWebhookBodyFormat.Xml,
                Parameters = new Dictionary<string, string>
                {
                    {"Param01", "Value01"},
                    {"Param02", "Value01"},
                }
            });

            result.ProvisioningTemplateWebhooks.Add(new PnP.Framework.Provisioning.Model.ProvisioningTemplateWebhook
            {
                Url = "https://my.url/func01",
                Method = ProvisioningTemplateWebhookMethod.POST,
                Async = true,
                Kind = ProvisioningTemplateWebhookKind.ProvisioningTemplateCompleted,
                BodyFormat = ProvisioningTemplateWebhookBodyFormat.FormUrlEncoded,
                Parameters = new Dictionary<string, string>
                {
                    {"Param01", "Value01"},
                    {"Param02", "Value01"},
                }
            });

            var serializer = new TargetSerializer();
            provider.SaveAs(result, TEST_OUT_FILE, serializer);

            var path = $"{provider.Connector.Parameters["ConnectionString"]}\\{provider.Connector.Parameters["Container"]}\\{TEST_OUT_FILE}";
            Assert.IsTrue(File.Exists(path)); path = Path.GetFullPath(path);
            var xml = XDocument.Load(path);
            var wrappedResult = XMLSerializer.Deserialize<TargetProvisioning>(xml);

            var template = wrappedResult.Templates[0].ProvisioningTemplate.First();

            Assert.AreEqual(2, template.ProvisioningTemplateWebhooks.Count());
            Assert.AreEqual("https://my.url/func01", template.ProvisioningTemplateWebhooks[0].Url);
            Assert.AreEqual(false, template.ProvisioningTemplateWebhooks[0].Async);
            Assert.AreEqual(ProvisioningWebhookBodyFormat.Xml, template.ProvisioningTemplateWebhooks[0].BodyFormat);
            Assert.AreEqual(ProvisioningWebhookMethod.GET, template.ProvisioningTemplateWebhooks[0].Method);
            Assert.AreEqual(ProvisioningWebhookMethod.POST, template.ProvisioningTemplateWebhooks[1].Method);
            Assert.AreEqual(ProvisioningWebhookKind.ProvisioningTemplateCompleted, template.ProvisioningTemplateWebhooks[1].Kind);
            Assert.AreEqual(ProvisioningWebhookKind.ObjectHandlerProvisioningStarted, template.ProvisioningTemplateWebhooks[0].Kind);
            Assert.AreEqual(2, template.ProvisioningTemplateWebhooks[0].Parameters.Count());
            Assert.AreEqual("Param01", template.ProvisioningTemplateWebhooks[0].Parameters[0].Key);
            Assert.AreEqual("Value01", template.ProvisioningTemplateWebhooks[0].Parameters[0].Value);
            Assert.AreEqual("Param01", template.ProvisioningTemplateWebhooks[1].Parameters[0].Key);
            Assert.AreEqual("Value01", template.ProvisioningTemplateWebhooks[1].Parameters[0].Value);
        }

        [TestMethod]
        [TestCategory(TEST_CATEGORY)]
        public void XMLSerializer_Deserialize_ProvisioningWebhook()
        {
            var provider = new XMLFileSystemTemplateProvider($@"{AppDomain.CurrentDomain.BaseDirectory}\..\..\..\Resources", "Templates");

            var hierarchy = provider.GetHierarchy(TEST_TEMPLATE);

            Assert.AreEqual(2, hierarchy.ProvisioningWebhooks.Count);
            Assert.IsTrue(hierarchy.ProvisioningWebhooks[0].Url.Contains("https://mywebhook.azurefunctions.net/function01"));
            Assert.AreEqual(ProvisioningTemplateWebhookMethod.GET, hierarchy.ProvisioningWebhooks[0].Method);
            Assert.AreEqual(ProvisioningTemplateWebhookKind.ProvisioningStarted, hierarchy.ProvisioningWebhooks[0].Kind);
            Assert.AreEqual(ProvisioningTemplateWebhookBodyFormat.Json, hierarchy.ProvisioningWebhooks[1].BodyFormat);
            Assert.AreEqual(true, hierarchy.ProvisioningWebhooks[1].Async);
            Assert.AreEqual(2, hierarchy.ProvisioningWebhooks[0].Parameters.Count);
            Assert.IsTrue(hierarchy.ProvisioningWebhooks[0].Parameters.ContainsKey("Param01"));
            Assert.AreEqual("Value01", hierarchy.ProvisioningWebhooks[0].Parameters["Param01"]);
            Assert.AreEqual("{sitecollection}", hierarchy.ProvisioningWebhooks[1].Parameters["Site"]);
        }

        [TestMethod]
        [TestCategory(TEST_CATEGORY)]
        public void XMLSerializer_Serialize_ProvisioningWebhook()
        {
            var provider = new XMLFileSystemTemplateProvider($@"{AppDomain.CurrentDomain.BaseDirectory}\..\..\..\Resources", "Templates");

            var result = new ProvisioningTemplate { ParentHierarchy = new ProvisioningHierarchy() };

            result.ParentHierarchy.ProvisioningWebhooks.Add(new PnP.Framework.Provisioning.Model.ProvisioningWebhook
            {
                Url = "https://my.url/func01",
                Method = ProvisioningTemplateWebhookMethod.GET,
                Async = false,
                Kind = ProvisioningTemplateWebhookKind.ObjectHandlerProvisioningStarted,
                BodyFormat = ProvisioningTemplateWebhookBodyFormat.Xml,
                Parameters = new Dictionary<string, string>
                {
                    {"Param01", "Value01"},
                    {"Param02", "Value01"},
                }
            });

            result.ParentHierarchy.ProvisioningWebhooks.Add(new PnP.Framework.Provisioning.Model.ProvisioningWebhook
            {
                Url = "https://my.url/func01",
                Method = ProvisioningTemplateWebhookMethod.POST,
                Async = true,
                Kind = ProvisioningTemplateWebhookKind.ProvisioningCompleted,
                BodyFormat = ProvisioningTemplateWebhookBodyFormat.FormUrlEncoded,
                Parameters = new Dictionary<string, string>
                {
                    {"Param01", "Value01"},
                    {"Param02", "Value01"},
                }
            });

            var serializer = new TargetSerializer();
            provider.SaveAs(result, TEST_OUT_FILE, serializer);

            var path = $"{provider.Connector.Parameters["ConnectionString"]}\\{provider.Connector.Parameters["Container"]}\\{TEST_OUT_FILE}";
            Assert.IsTrue(File.Exists(path)); path = Path.GetFullPath(path);
            var xml = XDocument.Load(path);
            var wrappedResult = XMLSerializer.Deserialize<TargetProvisioning>(xml);

            Assert.AreEqual(2, wrappedResult.ProvisioningWebhooks.Count());
            Assert.AreEqual("https://my.url/func01", wrappedResult.ProvisioningWebhooks[0].Url);
            Assert.AreEqual(false, wrappedResult.ProvisioningWebhooks[0].Async);
            Assert.AreEqual(ProvisioningWebhookBodyFormat.Xml, wrappedResult.ProvisioningWebhooks[0].BodyFormat);
            Assert.AreEqual(ProvisioningWebhookMethod.GET, wrappedResult.ProvisioningWebhooks[0].Method);
            Assert.AreEqual(ProvisioningWebhookMethod.POST, wrappedResult.ProvisioningWebhooks[1].Method);
            Assert.AreEqual(ProvisioningWebhookKind.ProvisioningCompleted, wrappedResult.ProvisioningWebhooks[1].Kind);
            Assert.AreEqual(ProvisioningWebhookKind.ObjectHandlerProvisioningStarted, wrappedResult.ProvisioningWebhooks[0].Kind);
            Assert.AreEqual(2, wrappedResult.ProvisioningWebhooks[0].Parameters.Count());
            Assert.AreEqual("Param01", wrappedResult.ProvisioningWebhooks[0].Parameters[0].Key);
            Assert.AreEqual("Value01", wrappedResult.ProvisioningWebhooks[0].Parameters[0].Value);
            Assert.AreEqual("Param01", wrappedResult.ProvisioningWebhooks[1].Parameters[0].Key);
            Assert.AreEqual("Value01", wrappedResult.ProvisioningWebhooks[1].Parameters[0].Value);
        }

        [TestMethod]
        [TestCategory(TEST_CATEGORY)]
        public void XMLSerializer_Deserialize_TeamTemplate()
        {
            var provider = new XMLFileSystemTemplateProvider($@"{AppDomain.CurrentDomain.BaseDirectory}\..\..\..\Resources", "Templates");

            var hierarchy = provider.GetHierarchy(TEST_TEMPLATE);
            var teamsTemplate = hierarchy.Teams.TeamTemplates;

            Assert.AreEqual(1, teamsTemplate.Count);
            Assert.AreEqual("Custom", teamsTemplate.First().Classification);
            Assert.AreEqual(TeamVisibility.Public, teamsTemplate.First().Visibility);
            Assert.IsTrue(teamsTemplate.First().JsonTemplate.Contains("here goes the team template JSON"));
            Assert.AreEqual("Sample Team from Template", teamsTemplate.First().Description);
            Assert.AreEqual("Team from Template 01", teamsTemplate.First().DisplayName);
        }

        [TestMethod]
        [TestCategory(TEST_CATEGORY)]
        public void XMLSerializer_Serialize_TeamTemplate()
        {
            var provider = new XMLFileSystemTemplateProvider($@"{AppDomain.CurrentDomain.BaseDirectory}\..\..\..\Resources", "Templates");

            var result = new ProvisioningTemplate { ParentHierarchy = new ProvisioningHierarchy() };

            result.ParentHierarchy.Teams.TeamTemplates.Add(new TeamTemplate
            {
                Classification = "MyClass",
                JsonTemplate = "{JSON template here}",
                Description = "Desc",
                Visibility = TeamVisibility.Private,
                DisplayName = "MyTemplate"
            });

            var serializer = new TargetSerializer();
            provider.SaveAs(result, TEST_OUT_FILE, serializer);

            var path = $"{provider.Connector.Parameters["ConnectionString"]}\\{provider.Connector.Parameters["Container"]}\\{TEST_OUT_FILE}";
            Assert.IsTrue(File.Exists(path)); path = Path.GetFullPath(path);
            var xml = XDocument.Load(path);
            var wrappedResult = XMLSerializer.Deserialize<TargetProvisioning>(xml);

            var teamTempaltes = wrappedResult.Teams.Items
                .Where(t => t is TargetNamespace.TeamTemplate).Cast<TargetNamespace.TeamTemplate>().ToList();

            Assert.AreEqual(1, teamTempaltes.Count);
            Assert.AreEqual("MyClass", teamTempaltes[0].Classification);
            Assert.AreEqual("Desc", teamTempaltes[0].Description);
            Assert.AreEqual("MyTemplate", teamTempaltes[0].DisplayName);
            Assert.AreEqual(BaseTeamVisibility.Private, teamTempaltes[0].Visibility);
            Assert.IsTrue(teamTempaltes[0].Text[0].Contains("JSON template here"));
        }

        [TestMethod]
        [TestCategory(TEST_CATEGORY)]
        public void XMLSerializer_Deserialize_Team()
        {
            var provider = new XMLFileSystemTemplateProvider($@"{AppDomain.CurrentDomain.BaseDirectory}\..\..\..\Resources", "Templates");

            var hierarchy = provider.GetHierarchy(TEST_TEMPLATE);
            var teams = hierarchy.Teams.Teams;

            Assert.AreEqual(2, teams.Count);

            // team common properties
            Assert.AreEqual("Sample Team 01", teams[0].DisplayName);
            Assert.AreEqual("This is just a sample Team 01", teams[0].Description);
            Assert.AreEqual("{o365groupid:GroupMailNickname}", teams[0].CloneFrom);
            Assert.AreEqual("{o365groupid:GroupMailNickname}", teams[1].GroupId);
            Assert.AreEqual("Private", teams[1].Classification);
            Assert.AreEqual(TeamSpecialization.EducationStandard, teams[1].Specialization);
            Assert.AreEqual(TeamVisibility.Public, teams[1].Visibility);
            Assert.AreEqual(false, teams[1].Archived);
            Assert.AreEqual("sample-team-02", teams[1].MailNickname);
            Assert.AreEqual("photo.jpg", teams[1].Photo);
            Assert.IsTrue(teams[0].HiddenGroupMembershipEnabled);

            // team security
            var security = teams[0].Security;
            Assert.AreEqual(true, security.AllowToAddGuests);
            Assert.AreEqual(false, security.ClearExistingMembers);
            Assert.AreEqual(true, security.ClearExistingOwners);
            Assert.AreEqual(2, security.Owners.Count);
            Assert.AreEqual("owner01@domain.onmicrosoft.com", security.Owners[0].UserPrincipalName);
            Assert.AreEqual(3, security.Members.Count);
            Assert.AreEqual("user01@domain.onmicrosoft.com", security.Members[0].UserPrincipalName);

            // team fun settings
            Assert.AreEqual(true, teams[1].FunSettings.AllowCustomMemes);
            Assert.AreEqual(true, teams[1].FunSettings.AllowGiphy);
            Assert.AreEqual(true, teams[1].FunSettings.AllowStickersAndMemes);
            Assert.AreEqual(TeamGiphyContentRating.Strict, teams[1].FunSettings.GiphyContentRating);

            // team guest settings
            Assert.AreEqual(true, teams[1].GuestSettings.AllowCreateUpdateChannels);
            Assert.AreEqual(false, teams[1].GuestSettings.AllowDeleteChannels);

            // team memebers settings
            Assert.AreEqual(false, teams[1].MemberSettings.AllowDeleteChannels);
            Assert.AreEqual(true, teams[1].MemberSettings.AllowAddRemoveApps);
            Assert.AreEqual(true, teams[1].MemberSettings.AllowCreateUpdateChannels);
            Assert.AreEqual(true, teams[1].MemberSettings.AllowCreateUpdateRemoveConnectors);
            Assert.AreEqual(false, teams[1].MemberSettings.AllowCreateUpdateRemoveTabs);
            Assert.AreEqual(true, teams[1].MemberSettings.AllowCreatePrivateChannels);

            // team messaging settings
            Assert.AreEqual(true, teams[1].MessagingSettings.AllowChannelMentions);
            Assert.AreEqual(false, teams[1].MessagingSettings.AllowOwnerDeleteMessages);
            Assert.AreEqual(true, teams[1].MessagingSettings.AllowTeamMentions);
            Assert.AreEqual(true, teams[1].MessagingSettings.AllowUserDeleteMessages);
            Assert.AreEqual(true, teams[1].MessagingSettings.AllowUserEditMessages);

            // team discovery settings
            Assert.AreEqual(false, teams[1].DiscoverySettings.ShowInTeamsSearchAndSuggestions);

            // team channels
            var channels = teams[1].Channels;
            Assert.AreEqual(3, channels.Count);
            Assert.AreEqual("12345", channels[0].ID);
            Assert.AreEqual("This is just a Sample Channel", channels[0].Description);
            Assert.AreEqual("Sample Channel 01", channels[0].DisplayName);
            Assert.AreEqual(true, channels[0].IsFavoriteByDefault);
            Assert.AreEqual(2, channels[0].Tabs.Count);
            Assert.AreEqual(true, channels[0].Tabs[0].Remove);
            Assert.AreEqual("67890", channels[0].Tabs[1].ID);
            Assert.AreEqual("My Tab 01", channels[0].Tabs[1].DisplayName);
            Assert.AreEqual("12345", channels[0].Tabs[1].TeamsAppId);
            Assert.AreEqual("https://www.contoso.com/Orders/2DCA2E6C7A10415CAF6B8AB6661B3154/tabView", channels[0].Tabs[1].Configuration.ContentUrl);
            Assert.AreEqual("2DCA2E6C7A10415CAF6B8AB6661B3154", channels[0].Tabs[1].Configuration.EntityId);
            Assert.AreEqual("https://www.contoso.com/Orders/2DCA2E6C7A10415CAF6B8AB6661B3154/uninstallTab", channels[0].Tabs[1].Configuration.RemoveUrl);
            Assert.AreEqual("https://www.contoso.com/Orders/2DCA2E6C7A10415CAF6B8AB6661B3154", channels[0].Tabs[1].Configuration.WebsiteUrl);
            Assert.IsTrue(channels[0].TabResources[0].TabResourceSettings.Contains("\"displayName\": \"Notebook name\""));
            Assert.AreEqual("{TeamsTabId:TabDisplayName}", channels[0].TabResources[0].TargetTabId);
            Assert.AreEqual(TabResourceType.Notebook, channels[0].TabResources[0].Type);
            Assert.AreEqual(1, channels[0].Messages.Count);
            Assert.IsTrue(channels[0].Messages[0].Message.Contains("Welcome to this channel"));
            Assert.AreEqual(MembershipType.Standard, channels[0].MembershipType);
            Assert.AreEqual(false, channels[0].Private);
            Assert.AreEqual(MembershipType.Shared, channels[1].MembershipType);
            Assert.AreEqual(false, channels[1].Private);
            Assert.AreEqual(MembershipType.Private, channels[2].MembershipType);
            Assert.AreEqual(true, channels[2].Private);
            Assert.IsTrue(channels[1].AllowNewMessageFromBots);
            Assert.IsTrue(channels[1].AllowNewMessageFromConnectors);
            Assert.AreEqual(ReplyRestriction.AuthorAndModerators, channels[1].ReplyRestriction);
            Assert.AreEqual(UserNewMessageRestriction.Moderators, channels[1].UserNewMessageRestriction);

            // team apps
            Assert.AreEqual(2, teams[1].Apps.Count);
            Assert.AreEqual("12345678-9abc-def0-123456789a", teams[1].Apps[0].AppId);
        }

        [TestMethod]
        [TestCategory(TEST_CATEGORY)]
        public void XMLSerializer_Serialize_Team()
        {
            var provider = new XMLFileSystemTemplateProvider($@"{AppDomain.CurrentDomain.BaseDirectory}\..\..\..\Resources", "Templates");

            var result = new ProvisioningTemplate { ParentHierarchy = new ProvisioningHierarchy() };

            result.ParentHierarchy.Teams.Teams.Add(new Team
            {
                DisplayName = "Sample Team 01",
                Description = "This is just a sample Team 01",
                CloneFrom = "{TeamId:GroupMailNickname}",
                Archived = false,
                MailNickname = "sample.group",
                GroupId = "12345",
                Classification = "Private",
                Specialization = TeamSpecialization.EducationClass,
                Visibility = TeamVisibility.Private,
                HiddenGroupMembershipEnabled = true,
                FunSettings = new TeamFunSettings
                {
                    AllowCustomMemes = true,
                    AllowGiphy = true,
                    GiphyContentRating = TeamGiphyContentRating.Moderate,
                    AllowStickersAndMemes = true
                },
                GuestSettings = new TeamGuestSettings
                {
                    AllowDeleteChannels = true,
                    AllowCreateUpdateChannels = true
                },
                MemberSettings = new TeamMemberSettings
                {
                    AllowDeleteChannels = true,
                    AllowCreateUpdateChannels = true,
                    AllowCreateUpdateRemoveTabs = true,
                    AllowCreateUpdateRemoveConnectors = true,
                    AllowAddRemoveApps = true,
                    AllowCreatePrivateChannels = true,
                },
                MessagingSettings = new TeamMessagingSettings
                {
                    AllowChannelMentions = true,
                    AllowTeamMentions = true,
                    AllowUserEditMessages = true,
                    AllowOwnerDeleteMessages = true,
                    AllowUserDeleteMessages = true
                },
                DiscoverySettings = new TeamDiscoverySettings
                {
                    ShowInTeamsSearchAndSuggestions = false,
                },
                Security = new PnP.Framework.Provisioning.Model.Teams.TeamSecurity
                {
                    AllowToAddGuests = true,
                    ClearExistingMembers = true,
                    ClearExistingOwners = true,
                    Members = {
                        new TeamSecurityUser
                        {
                            UserPrincipalName = "user01@domain.onmicrosoft.com"
                        },
                        new TeamSecurityUser
                        {
                            UserPrincipalName = "user02@domain.onmicrosoft.com"
                        }
                    },
                    Owners =
                    {
                        new TeamSecurityUser
                        {
                            UserPrincipalName = "owner01@domain.onmicrosoft.com"
                        },
                        new TeamSecurityUser
                        {
                            UserPrincipalName = "owner02@domain.onmicrosoft.com"
                        }
                    }
                },
                Apps =
                {
                    new TeamAppInstance
                    {
                        AppId = "12345678-9abc-def0-123456789a"
                    }
                },
                Channels =
                {
                    new PnP.Framework.Provisioning.Model.Teams.TeamChannel
                    {
                        ID = "12345",
                        Description = "This is just a Sample Channel",
                        DisplayName = "Sample Channel 01",
                        IsFavoriteByDefault = true,
                        // Private = true,
                        MembershipType = MembershipType.Shared,
                        AllowNewMessageFromBots = true,
                        AllowNewMessageFromConnectors = true,
                        ReplyRestriction = ReplyRestriction.AuthorAndModerators,
                        UserNewMessageRestriction = UserNewMessageRestriction.Moderators,
                        Tabs =
                        {
                            new TeamTab
                            {
                                DisplayName = "Wiki",
                                TeamsAppId = "com.microsoft.teamspace.tab.wiki",
                                ID = "{TeamsTabId:Wiki}",
                                Remove = true,
                            },
                            new TeamTab
                            {
                                DisplayName = "My Tab 01",
                                TeamsAppId = "12345",
                                ID = "67890",
                                Configuration = new TeamTabConfiguration
                                {
                                    ContentUrl = "https://www.contoso.com/Orders/2DCA2E6C7A10415CAF6B8AB6661B3154/tabView",
                                    WebsiteUrl = "https://www.contoso.com/Orders/2DCA2E6C7A10415CAF6B8AB6661B3154",
                                    EntityId = "2DCA2E6C7A10415CAF6B8AB6661B3154",
                                    RemoveUrl = "https://www.contoso.com/Orders/2DCA2E6C7A10415CAF6B8AB6661B3154/uninstallTab"
                                }
                            }
                        },
                        TabResources =
                        {
                            new TeamTabResource
                            {
                                TabResourceSettings = "{\"displayName\": \"Notebook name\"}",
                                TargetTabId = "{TeamsTabId:TabDisplayName}",
                                Type = TabResourceType.Planner
                            }
                        },
                        Messages =
                        {
                            new TeamChannelMessage
                            {
                                Message = "{\"body\": \"message\"}"
                            }
                        }
                    }
                }
            });

            var serializer = new TargetSerializer();
            provider.SaveAs(result, TEST_OUT_FILE, serializer);

            var path = $"{provider.Connector.Parameters["ConnectionString"]}\\{provider.Connector.Parameters["Container"]}\\{TEST_OUT_FILE}";
            Assert.IsTrue(File.Exists(path)); path = Path.GetFullPath(path);
            var xml = XDocument.Load(path);
            var wrappedResult = XMLSerializer.Deserialize<TargetProvisioning>(xml);

            var teams = wrappedResult.Teams.Items
                .Where(t => t is TeamWithSettings).Cast<TeamWithSettings>().ToList();

            Assert.AreEqual(1, teams.Count);
            var team = teams[0];

            // team common properties
            Assert.AreEqual("Sample Team 01", team.DisplayName);
            Assert.AreEqual("This is just a sample Team 01", team.Description);
            Assert.AreEqual("{TeamId:GroupMailNickname}", team.CloneFrom);
            Assert.AreEqual("12345", team.GroupId);
            Assert.AreEqual("Private", team.Classification);
            Assert.AreEqual(TeamWithSettingsSpecialization.EducationClass, team.Specialization);
            Assert.AreEqual(BaseTeamVisibility.Private, team.Visibility);
            Assert.AreEqual(false, team.Archived);
            Assert.AreEqual("sample.group", team.MailNickname);
            Assert.AreEqual(true, team.HiddenGroupMembershipEnabled);

            // team security
            var security = team.Security;
            Assert.AreEqual(true, security.AllowToAddGuests);
            Assert.AreEqual(true, security.Members.ClearExistingItems);
            Assert.AreEqual(true, security.Owners.ClearExistingItems);
            Assert.AreEqual(2, security.Owners.User.Count());
            Assert.AreEqual("owner01@domain.onmicrosoft.com", security.Owners.User[0].UserPrincipalName);
            Assert.AreEqual(2, security.Members.User.Count());
            Assert.AreEqual("user01@domain.onmicrosoft.com", security.Members.User[0].UserPrincipalName);

            // team fun settings
            Assert.AreEqual(true, team.FunSettings.AllowCustomMemes);
            Assert.AreEqual(true, team.FunSettings.AllowGiphy);
            Assert.AreEqual(true, team.FunSettings.AllowStickersAndMemes);
            Assert.AreEqual(TeamWithSettingsFunSettingsGiphyContentRating.Moderate, team.FunSettings.GiphyContentRating);

            // team guest settings
            Assert.AreEqual(true, team.GuestSettings.AllowCreateUpdateChannels);
            Assert.AreEqual(true, team.GuestSettings.AllowDeleteChannels);

            // team members settings
            Assert.AreEqual(true, team.MembersSettings.AllowDeleteChannels);
            Assert.AreEqual(true, team.MembersSettings.AllowAddRemoveApps);
            Assert.AreEqual(true, team.MembersSettings.AllowCreateUpdateChannels);
            Assert.AreEqual(true, team.MembersSettings.AllowCreateUpdateRemoveConnectors);
            Assert.AreEqual(true, team.MembersSettings.AllowCreateUpdateRemoveTabs);
            Assert.AreEqual(true, team.MembersSettings.AllowCreatePrivateChannels);

            // team messaging settings
            Assert.AreEqual(true, team.MessagingSettings.AllowChannelMentions);
            Assert.AreEqual(true, team.MessagingSettings.AllowOwnerDeleteMessages);
            Assert.AreEqual(true, team.MessagingSettings.AllowTeamMentions);
            Assert.AreEqual(true, team.MessagingSettings.AllowUserDeleteMessages);
            Assert.AreEqual(true, team.MessagingSettings.AllowUserEditMessages);

            // team discovery settings
            Assert.AreEqual(false, team.DiscoverySettings.ShowInTeamsSearchAndSuggestions);

            // team channels
            var channels = team.Channels;
            Assert.AreEqual(1, channels.Count());
            Assert.AreEqual("12345", channels[0].ID);
            Assert.AreEqual("This is just a Sample Channel", channels[0].Description);
            Assert.AreEqual("Sample Channel 01", channels[0].DisplayName);
            Assert.AreEqual(true, channels[0].IsFavoriteByDefault);
            Assert.AreEqual(2, channels[0].Tabs.Count());
            Assert.AreEqual(true, channels[0].Tabs[0].Remove);
            Assert.AreEqual("67890", channels[0].Tabs[1].ID);
            Assert.AreEqual("My Tab 01", channels[0].Tabs[1].DisplayName);
            Assert.AreEqual("12345", channels[0].Tabs[1].TeamsAppId);
            Assert.AreEqual("https://www.contoso.com/Orders/2DCA2E6C7A10415CAF6B8AB6661B3154/tabView", channels[0].Tabs[1].Configuration.ContentUrl);
            Assert.AreEqual("2DCA2E6C7A10415CAF6B8AB6661B3154", channels[0].Tabs[1].Configuration.EntityId);
            Assert.AreEqual("https://www.contoso.com/Orders/2DCA2E6C7A10415CAF6B8AB6661B3154/uninstallTab", channels[0].Tabs[1].Configuration.RemoveUrl);
            Assert.AreEqual("https://www.contoso.com/Orders/2DCA2E6C7A10415CAF6B8AB6661B3154", channels[0].Tabs[1].Configuration.WebsiteUrl);
            Assert.IsTrue(channels[0].TabResources[0].TabResourceSettings.Contains("\"displayName\": \"Notebook name\""));
            Assert.AreEqual("{TeamsTabId:TabDisplayName}", channels[0].TabResources[0].TargetTabId);
            Assert.AreEqual(TeamTabResourcesTabResourceType.Planner, channels[0].TabResources[0].Type);
            Assert.AreEqual(TeamChannelMembershipType.Shared, channels[0].MembershipType);
            Assert.IsTrue(channels[0].AllowNewMessageFromBots);
            Assert.IsTrue(channels[0].AllowNewMessageFromConnectors);
            Assert.AreEqual(TeamChannelReplyRestriction.AuthorAndModerators, channels[0].ReplyRestriction);
            Assert.AreEqual(TeamChannelUserNewMessageRestriction.Moderators, channels[0].UserNewMessageRestriction);

            // team apps
            Assert.AreEqual(1, team.Apps.Count());
            Assert.AreEqual("12345678-9abc-def0-123456789a", team.Apps[0].AppId);
        }

        [TestMethod]
        [TestCategory(TEST_CATEGORY)]
        public void XMLSerializer_Deserialize_TeamApps()
        {
            var provider = new XMLFileSystemTemplateProvider($@"{AppDomain.CurrentDomain.BaseDirectory}\..\..\..\Resources", "Templates");

            var hierarchy = provider.GetHierarchy(TEST_TEMPLATE);
            var apps = hierarchy.Teams.Apps;

            Assert.AreEqual(2, apps.Count);
            Assert.AreEqual("APP001", apps[0].AppId);
            Assert.AreEqual("./custom-app-01.json", apps[0].PackageUrl);
        }

        [TestMethod]
        [TestCategory(TEST_CATEGORY)]
        public void XMLSerializer_Serialize_TeamApps()
        {
            var provider = new XMLFileSystemTemplateProvider($@"{AppDomain.CurrentDomain.BaseDirectory}\..\..\..\Resources", "Templates");

            var result = new ProvisioningTemplate { ParentHierarchy = new ProvisioningHierarchy() };

            result.ParentHierarchy.Teams.Apps.Add(new TeamApp
            {
                AppId = "APP001",
                PackageUrl = "./custom-app-02.zip"
            });

            var serializer = new TargetSerializer();
            provider.SaveAs(result, TEST_OUT_FILE, serializer);

            var path = $"{provider.Connector.Parameters["ConnectionString"]}\\{provider.Connector.Parameters["Container"]}\\{TEST_OUT_FILE}";
            Assert.IsTrue(File.Exists(path)); path = Path.GetFullPath(path);
            var xml = XDocument.Load(path);
            var wrappedResult = XMLSerializer.Deserialize<TargetProvisioning>(xml);

            var teamApps = wrappedResult.Teams.Apps;

            Assert.AreEqual(1, teamApps.Count());
            Assert.AreEqual("APP001", teamApps[0].AppId);
            Assert.AreEqual("./custom-app-02.zip", teamApps[0].PackageUrl);
        }

        [TestMethod]
        [TestCategory(TEST_CATEGORY)]
        public void XMLSerializer_Deserialize_AzureAD()
        {
            var provider = new XMLFileSystemTemplateProvider($@"{AppDomain.CurrentDomain.BaseDirectory}\..\..\..\Resources", "Templates");

            var hierarchy = provider.GetHierarchy(TEST_TEMPLATE);
            var users = hierarchy.AzureActiveDirectory.Users;

            Assert.AreEqual(2, users.Count);
            Assert.AreEqual("John White", users[0].DisplayName);
            Assert.AreEqual(true, users[0].AccountEnabled);
            Assert.AreEqual("john.white", users[0].MailNickname);
            Assert.AreEqual("john.white@{parameter:domain}.onmicrosoft.com", users[0].UserPrincipalName);
            Assert.AreEqual("DisablePasswordExpiration,DisableStrongPassword", users[0].PasswordPolicies);
            Assert.AreEqual("photo.jpg", users[0].ProfilePhoto);
            Assert.AreEqual("John", users[0].GivenName);
            Assert.AreEqual("White", users[0].Surname);
            Assert.AreEqual("Senior Partner", users[0].JobTitle);
            Assert.AreEqual("+1-601-123456", users[0].MobilePhone);
            Assert.AreEqual("Seattle, WA", users[0].OfficeLocation);
            Assert.AreEqual("US", users[0].UsageLocation);
            Assert.AreEqual("en-US", users[0].PreferredLanguage);

            var passWord = new SecureString();

            foreach (char c in "Pass@w0rd") passWord.AppendChar(c);

            Assert.AreEqual(true, users[0].PasswordProfile.ForceChangePasswordNextSignIn);
            Assert.AreEqual(false, users[0].PasswordProfile.ForceChangePasswordNextSignInWithMfa);
            Assert.IsFalse(users[0].PasswordProfile.Password == null);
            Assert.AreEqual(2, users[0].Licenses.Count);
            Assert.AreEqual("6fd2c87f-b296-42f0-b197-1e91e994b900", users[0].Licenses[0].SkuId);
            Assert.AreEqual("5136a095-5cf0-4aff-bec3-e84448b38ea5", users[0].Licenses[0].DisabledPlans[0]);
        }

        [TestMethod]
        [TestCategory(TEST_CATEGORY)]
        public void XMLSerializer_Serialize_AzureAD()
        {
            var provider = new XMLFileSystemTemplateProvider($@"{AppDomain.CurrentDomain.BaseDirectory}\..\..\..\Resources", "Templates");

            var passWord = new SecureString();
            foreach (char c in "Pass@w0rd") passWord.AppendChar(c);

            var result = new ProvisioningTemplate { ParentHierarchy = new ProvisioningHierarchy() };
            result.ParentHierarchy.AzureActiveDirectory.Users.Add(new PnP.Framework.Provisioning.Model.AzureActiveDirectory.User
            {
                AccountEnabled = true,
                DisplayName = "John White",
                MailNickname = "john.white",
                UserPrincipalName = "john.white@{parameter:domain}.onmicrosoft.com",
                PasswordPolicies = "Policy1",
                ProfilePhoto = "photo.jpg",
                GivenName = "John",
                Surname = "White",
                JobTitle = "Senior Partner",
                MobilePhone = "+1-601-123456",
                OfficeLocation = "Seattle, WA",
                PreferredLanguage = "en-US",
                PasswordProfile = new PasswordProfile
                {
                    ForceChangePasswordNextSignIn = true,
                    ForceChangePasswordNextSignInWithMfa = true,
                    Password = passWord
                },
                Licenses =
                {
                    new UserLicense
                    {
                        SkuId = "26d45bd9-adf1-46cd-a9e1-51e9a5524128",
                        DisabledPlans = new []{ "e212cbc7-0961-4c40-9825-01117710dcb1", "3fb82609-8c27-4f7b-bd51-30634711ee67", "b1188c4c-1b36-4018-b48b-ee07604f6feb" }
                    },
                    new UserLicense
                    {
                        SkuId = "26d45bd9-adf1-46cd-a9e1-51e9a5524128"
                    }
                }
            });


            var serializer = new TargetSerializer();
            provider.SaveAs(result, TEST_OUT_FILE, serializer);

            var path = $"{provider.Connector.Parameters["ConnectionString"]}\\{provider.Connector.Parameters["Container"]}\\{TEST_OUT_FILE}";
            Assert.IsTrue(File.Exists(path)); path = Path.GetFullPath(path);
            var xml = XDocument.Load(path);
            var wrappedResult = XMLSerializer.Deserialize<TargetProvisioning>(xml);
            var users = wrappedResult.AzureActiveDirectory.Users;

            Assert.AreEqual(1, users.Count());
            Assert.AreEqual("John White", users[0].DisplayName);
            Assert.AreEqual(true, users[0].AccountEnabled);
            Assert.AreEqual("john.white", users[0].MailNickname);
            Assert.AreEqual("john.white@{parameter:domain}.onmicrosoft.com", users[0].UserPrincipalName);
            Assert.AreEqual("Policy1", users[0].PasswordPolicies);
            Assert.AreEqual("photo.jpg", users[0].ProfilePhoto);
            Assert.AreEqual("John", users[0].GivenName);
            Assert.AreEqual("White", users[0].Surname);
            Assert.AreEqual("Senior Partner", users[0].JobTitle);
            Assert.AreEqual("+1-601-123456", users[0].MobilePhone);
            Assert.AreEqual("Seattle, WA", users[0].OfficeLocation);
            Assert.AreEqual("en-US", users[0].PreferredLanguage);

            Assert.AreEqual(true, users[0].PasswordProfile.ForceChangePasswordNextSignIn);
            Assert.AreEqual(true, users[0].PasswordProfile.ForceChangePasswordNextSignInWithMfa);
            Assert.IsFalse(users[0].PasswordProfile.Password == null);
            Assert.AreEqual(2, users[0].Licenses.Count());
            Assert.AreEqual("26d45bd9-adf1-46cd-a9e1-51e9a5524128", users[0].Licenses[0].SkuId);
            Assert.AreEqual("e212cbc7-0961-4c40-9825-01117710dcb1", users[0].Licenses[0].DisabledPlans[0]);
        }

        [TestMethod]
        [TestCategory(TEST_CATEGORY)]
        public void XMLSerializer_Deserialize_Drive()
        {
            var provider = new XMLFileSystemTemplateProvider($@"{AppDomain.CurrentDomain.BaseDirectory}\..\..\..\Resources", "Templates");

            var hierarchy = provider.GetHierarchy(TEST_TEMPLATE);
            var driveRoots = hierarchy.Drive.DriveRoots;

            Assert.AreEqual(2, driveRoots.Count);
            Assert.AreEqual("/users/jim.black@contoso.onmicrosoft.com/drive", driveRoots[0].DriveUrl);
            Assert.AreEqual(2, driveRoots[0].RootFolder.DriveFiles.Count);
            Assert.AreEqual(2, driveRoots[0].RootFolder.DriveFolders.Count);
            Assert.AreEqual("Projects", driveRoots[0].RootFolder.DriveFolders[0].Name);
            Assert.AreEqual("OneDrive/Jim.Black/Projects", driveRoots[0].RootFolder.DriveFolders[0].Src);
            Assert.AreEqual("ExpenseReports", driveRoots[0].RootFolder.DriveFolders[1].Name);
            Assert.AreEqual(2, driveRoots[0].RootFolder.DriveFolders[1].DriveFolders.Count);
            Assert.AreEqual("FY2018", driveRoots[0].RootFolder.DriveFolders[1].DriveFolders[0].Name);
            Assert.AreEqual(12, driveRoots[0].RootFolder.DriveFolders[1].DriveFolders[0].DriveFiles.Count);
            Assert.AreEqual("Expense-Report-Jan2018.xlsx", driveRoots[0].RootFolder.DriveFolders[1].DriveFolders[0].DriveFiles[0].Name);
            Assert.AreEqual("OneDrive/Jim.Black/ExpenseReports/FY2018/Expense-Report-Jan2018.xlsx", driveRoots[0].RootFolder.DriveFolders[1].DriveFolders[0].DriveFiles[0].Src);
            Assert.AreEqual(false, driveRoots[0].RootFolder.DriveFolders[1].DriveFolders[0].DriveFiles[0].Overwrite);
            Assert.AreEqual("FY2019", driveRoots[0].RootFolder.DriveFolders[1].DriveFolders[1].Name);
            Assert.AreEqual(9, driveRoots[0].RootFolder.DriveFolders[1].DriveFolders[1].DriveFiles.Count);
            Assert.AreEqual("MyFavoriteSlides.pptx", driveRoots[0].RootFolder.DriveFiles[0].Name);
            Assert.AreEqual("OneDrive/Jim.Black/MyFavoriteSlides.pptx", driveRoots[0].RootFolder.DriveFiles[0].Src);
            Assert.AreEqual("/users/john.white@{parameter:O365TenantName}.onmicrosoft.com/drive", driveRoots[1].DriveUrl);
        }

        [TestMethod]
        [TestCategory(TEST_CATEGORY)]
        public void XMLSerializer_Serialize_Drive()
        {
            var provider = new XMLFileSystemTemplateProvider($@"{AppDomain.CurrentDomain.BaseDirectory}\..\..\..\Resources", "Templates");

            #region Prepare source model

            var result = new ProvisioningTemplate { ParentHierarchy = new ProvisioningHierarchy() };
            var driveRootOne = new PnP.Framework.Provisioning.Model.Drive.DriveRoot
            {
                DriveUrl = "/users/jim.black@contoso.onmicrosoft.com/drive",
                RootFolder = new PnP.Framework.Provisioning.Model.Drive.DriveRootFolder()
            };
            driveRootOne.RootFolder.DriveFiles.AddRange(new PnP.Framework.Provisioning.Model.Drive.DriveFile[]
            {
                new PnP.Framework.Provisioning.Model.Drive.DriveFile
                {
                    Name = "MyFavoriteSlides.pptx",
                    Src = "OneDrive/Jim.Black/MyFavoriteSlides.pptx",
                    Overwrite = false,
                },
                new PnP.Framework.Provisioning.Model.Drive.DriveFile
                {
                    Name = "Jim-Black-Resume.docx",
                    Src = "OneDrive/Jim.Black/Jim-Black-Resume.docx",
                    Overwrite = false,
                },
            });
            driveRootOne.RootFolder.DriveFolders.AddRange(new PnP.Framework.Provisioning.Model.Drive.DriveFolder[]
            {
                new PnP.Framework.Provisioning.Model.Drive.DriveFolder
                {
                    Name = "Projects",
                    Src = "OneDrive/Jim.Black/Projects",
                    Overwrite = false,
                    IncludedExtensions = "*.pdf",
                    ExcludedExtensions = "*.xlsx,*.pptx",
                    Recursive = true
                },
                new PnP.Framework.Provisioning.Model.Drive.DriveFolder
                {
                    Name = "ExpenseReports",
                    Overwrite = true,
                },
            });
            driveRootOne.RootFolder.DriveFolders[1].DriveFolders.Add(
                new PnP.Framework.Provisioning.Model.Drive.DriveFolder
                {
                    Name = "FY2018",
                    Overwrite = true,
                });
            driveRootOne.RootFolder.DriveFolders[1].DriveFolders[0].DriveFiles.AddRange(
                from f in Enumerable.Range(1, 12)
                select new PnP.Framework.Provisioning.Model.Drive.DriveFile
                {
                    Name = $"Expense-Report-{String.Format("{0:00}", f)}-2018.xlsx",
                    Src = $"OneDrive/Jim.Black/ExpenseReports/FY2018/Expense-Report-{String.Format("{0:00}", f)}-2018.xlsx",
                    Overwrite = false,
                }
                );
            driveRootOne.RootFolder.DriveFolders[1].DriveFolders.Add(
                new PnP.Framework.Provisioning.Model.Drive.DriveFolder
                {
                    Name = "FY2019",
                    Overwrite = true,
                });
            driveRootOne.RootFolder.DriveFolders[1].DriveFolders[1].DriveFiles.AddRange(
                from f in Enumerable.Range(1, 9)
                select new PnP.Framework.Provisioning.Model.Drive.DriveFile
                {
                    Name = $"Expense-Report-{String.Format("{0:00}", f)}-2019.xlsx",
                    Src = $"OneDrive/Jim.Black/ExpenseReports/FY2018/Expense-Report-{String.Format("{0:00}", f)}-2019.xlsx",
                    Overwrite = true,
                }
                );
            result.ParentHierarchy.Drive.DriveRoots.Add(driveRootOne);

            var driveRootTwo = new PnP.Framework.Provisioning.Model.Drive.DriveRoot
            {
                DriveUrl = "/users/john.white@{parameter:O365TenantName}.onmicrosoft.com/drive",
                RootFolder = new PnP.Framework.Provisioning.Model.Drive.DriveRootFolder()
            };
            driveRootTwo.RootFolder.DriveFiles.AddRange(new PnP.Framework.Provisioning.Model.Drive.DriveFile[]
            {
                new PnP.Framework.Provisioning.Model.Drive.DriveFile
                {
                    Name = "John-White-Resume.docx",
                    Src = "OneDrive/John.White/JWResume.docx",
                    Overwrite = true,
                },
            });
            driveRootTwo.RootFolder.DriveFolders.AddRange(new PnP.Framework.Provisioning.Model.Drive.DriveFolder[]
            {
                new PnP.Framework.Provisioning.Model.Drive.DriveFolder
                {
                    Name = "Documents",
                    Src = "OneDrive/John.White/Documents",
                    Overwrite = false,
                    IncludedExtensions = "*.pdf",
                    ExcludedExtensions = "*.xlsx,*.pptx",
                    Recursive = true
                },
            });
            result.ParentHierarchy.Drive.DriveRoots.Add(driveRootTwo);

            #endregion

            var serializer = new TargetSerializer();
            provider.SaveAs(result, TEST_OUT_FILE, serializer);

            var path = $"{provider.Connector.Parameters["ConnectionString"]}\\{provider.Connector.Parameters["Container"]}\\{TEST_OUT_FILE}";
            Assert.IsTrue(File.Exists(path)); path = Path.GetFullPath(path);
            var xml = XDocument.Load(path);
            var wrappedResult = XMLSerializer.Deserialize<TargetProvisioning>(xml);
            var drives = wrappedResult.Drive;

            Assert.AreEqual(2, drives.Count());
            Assert.AreEqual("/users/jim.black@contoso.onmicrosoft.com/drive", drives[0].DriveUrl);
            Assert.AreEqual(4, drives[0].DriveItems.Length);
            Assert.AreEqual("/users/john.white@{parameter:O365TenantName}.onmicrosoft.com/drive", drives[1].DriveUrl);
            Assert.AreEqual(2, drives[1].DriveItems.Length);

            Assert.AreEqual("ExpenseReports", ((DriveFolder)drives[0].DriveItems[1]).Name);
            Assert.AreEqual(2, ((DriveFolder)drives[0].DriveItems[1]).Items.Length);
            Assert.AreEqual("FY2018", ((DriveFolder)((DriveFolder)drives[0].DriveItems[1]).Items[0]).Name);
            Assert.AreEqual(12, ((DriveFolder)((DriveFolder)drives[0].DriveItems[1]).Items[0]).Items.Length);
            Assert.AreEqual("FY2019", ((DriveFolder)((DriveFolder)drives[0].DriveItems[1]).Items[1]).Name);
            Assert.AreEqual(9, ((DriveFolder)((DriveFolder)drives[0].DriveItems[1]).Items[1]).Items.Length);
            Assert.AreEqual("Expense-Report-01-2018.xlsx", ((DriveFile)((DriveFolder)((DriveFolder)drives[0].DriveItems[1]).Items[0]).Items[0]).Name);
            Assert.AreEqual("Expense-Report-02-2018.xlsx", ((DriveFile)((DriveFolder)((DriveFolder)drives[0].DriveItems[1]).Items[0]).Items[1]).Name);
            Assert.AreEqual("Expense-Report-12-2018.xlsx", ((DriveFile)((DriveFolder)((DriveFolder)drives[0].DriveItems[1]).Items[0]).Items[11]).Name);
        }

        [TestMethod]
        [TestCategory(TEST_CATEGORY)]
        public void XMLSerializer_Deserialize_Tenant_AppCatalog()
        {
            var provider = new XMLFileSystemTemplateProvider($@"{AppDomain.CurrentDomain.BaseDirectory}\..\..\..\Resources", "Templates");

            var hierarchy = provider.GetHierarchy(TEST_TEMPLATE);
            var appCatalog = hierarchy.Tenant.AppCatalog;

            Assert.AreEqual(3, appCatalog.Packages.Count);
            Assert.AreEqual("solution\\spfx-discuss-now.sppkg", appCatalog.Packages[0].Src);
            Assert.AreEqual(PackageAction.UploadAndPublish, appCatalog.Packages[0].Action);
            Assert.AreEqual(true, appCatalog.Packages[0].Overwrite);
            Assert.AreEqual(true, appCatalog.Packages[0].SkipFeatureDeployment);
            Assert.AreEqual("d0816f0a-fda4-4a98-8e61-1bbe1f2b5b27", appCatalog.Packages[1].PackageId);
            Assert.AreEqual(PackageAction.Publish, appCatalog.Packages[1].Action);
        }

        [TestMethod]
        [TestCategory(TEST_CATEGORY)]
        public void XMLSerializer_Serialize_Tenant_AppCatalog()
        {
            var provider = new XMLFileSystemTemplateProvider($@"{AppDomain.CurrentDomain.BaseDirectory}\..\..\..\Resources", "Templates");

            var passWord = new SecureString();
            foreach (char c in "Pass@w0rd") passWord.AppendChar(c);

            var result = new ProvisioningTemplate { ParentHierarchy = new ProvisioningHierarchy() };
            result.ParentHierarchy.Tenant = new ProvisioningTenant(result.ApplicationLifecycleManagement.AppCatalog, new PnP.Framework.Provisioning.Model.ContentDeliveryNetwork());

            result.Tenant.AppCatalog.Packages.Add(new Package
            {
                Action = PackageAction.Publish,
                Src = "solution\\spfx-discuss-now.sppkg",
                SkipFeatureDeployment = true,
                Overwrite = true
            });
            result.Tenant.AppCatalog.Packages.Add(new Package
            {
                Action = PackageAction.Upload,
                SkipFeatureDeployment = true,
                Overwrite = true,
                PackageId = "d0816f0a-fda4-4a98-8e61-1bbe1f2b5b27"
            });

            var serializer = new TargetSerializer();
            provider.SaveAs(result, TEST_OUT_FILE, serializer);

            var path = $"{provider.Connector.Parameters["ConnectionString"]}\\{provider.Connector.Parameters["Container"]}\\{TEST_OUT_FILE}";
            Assert.IsTrue(File.Exists(path)); path = Path.GetFullPath(path);
            var xml = XDocument.Load(path);
            var wrappedResult = XMLSerializer.Deserialize<TargetProvisioning>(xml);
            var packages = wrappedResult.Tenant.AppCatalog;

            Assert.AreEqual(2, packages.Count());
            Assert.AreEqual("solution\\spfx-discuss-now.sppkg", packages[0].Src);
            Assert.AreEqual(AppCatalogPackageAction.Publish, packages[0].Action);
            Assert.AreEqual(true, packages[0].Overwrite);
            Assert.AreEqual(true, packages[0].SkipFeatureDeployment);
            Assert.AreEqual("d0816f0a-fda4-4a98-8e61-1bbe1f2b5b27", packages[1].PackageId);
            Assert.AreEqual(AppCatalogPackageAction.Upload, packages[1].Action);
        }

        [TestMethod]
        [TestCategory(TEST_CATEGORY)]
        public void XMLSerializer_Deserialize_Tenant_WebApiPermissions()
        {
            var provider = new XMLFileSystemTemplateProvider($@"{AppDomain.CurrentDomain.BaseDirectory}\..\..\..\Resources", "Templates");

            var hierarchy = provider.GetHierarchy(TEST_TEMPLATE);
            var apiPermission = hierarchy.Tenant.WebApiPermissions;

            Assert.AreEqual("Microsoft.Graph", apiPermission[0].Resource);
            Assert.AreEqual("User.ReadBasic.All", apiPermission[0].Scope);
        }

        [TestMethod]
        [TestCategory(TEST_CATEGORY)]
        public void XMLSerializer_Serialize_Tenant_WebApiPermissions()
        {
            var provider = new XMLFileSystemTemplateProvider($@"{AppDomain.CurrentDomain.BaseDirectory}\..\..\..\Resources", "Templates");

            var passWord = new SecureString();
            foreach (char c in "Pass@w0rd") passWord.AppendChar(c);

            var result = new ProvisioningTemplate { ParentHierarchy = new ProvisioningHierarchy() };
            result.ParentHierarchy.Tenant = new ProvisioningTenant(result.ApplicationLifecycleManagement.AppCatalog, new PnP.Framework.Provisioning.Model.ContentDeliveryNetwork());

            result.Tenant.WebApiPermissions.Add(
                new WebApiPermission
                {
                    Resource = "Microsoft.Graph",
                    Scope = "User.ReadBasic.All"
                }
            );

            var serializer = new TargetSerializer();
            provider.SaveAs(result, TEST_OUT_FILE, serializer);

            var path = $"{provider.Connector.Parameters["ConnectionString"]}\\{provider.Connector.Parameters["Container"]}\\{TEST_OUT_FILE}";
            Assert.IsTrue(File.Exists(path)); path = Path.GetFullPath(path);
            var xml = XDocument.Load(path);
            var wrappedResult = XMLSerializer.Deserialize<TargetProvisioning>(xml);
            var apiPermissions = wrappedResult.Tenant.WebApiPermissions;

            Assert.AreEqual("Microsoft.Graph", apiPermissions[0].Resource);
            Assert.AreEqual("User.ReadBasic.All", apiPermissions[0].Scope);
        }

        [TestMethod]
        [TestCategory(TEST_CATEGORY)]
        public void XMLSerializer_Deserialize_Tenant_ContentDeliveryNetwork()
        {
            var provider = new XMLFileSystemTemplateProvider($@"{AppDomain.CurrentDomain.BaseDirectory}\..\..\..\Resources", "Templates");

            var hierarchy = provider.GetHierarchy(TEST_TEMPLATE);
            var cdn = hierarchy.Tenant.ContentDeliveryNetwork;

            Assert.AreEqual(true, cdn.PublicCdn.Enabled);
            Assert.AreEqual("JS,CSS", cdn.PublicCdn.ExcludeIfNoScriptDisabled);
            Assert.AreEqual("HBI,GDPR", cdn.PublicCdn.ExcludeRestrictedSiteClassifications);
            Assert.AreEqual("PDF,XML,JPG,JS,PNG", cdn.PublicCdn.IncludeFileExtensions);
            Assert.AreEqual(true, cdn.PublicCdn.NoDefaultOrigins);
            Assert.AreEqual(OriginAction.Add, cdn.PublicCdn.Origins[0].Action);
            Assert.AreEqual("sites/CDN/CDNFiles", cdn.PublicCdn.Origins[0].Url);

            Assert.AreEqual(true, cdn.PrivateCdn.Enabled);
            Assert.AreEqual("HIB,GDPR", cdn.PrivateCdn.ExcludeRestrictedSiteClassifications);
            Assert.AreEqual("PDF,XML,JPG,JS,PNG", cdn.PrivateCdn.IncludeFileExtensions);
            Assert.AreEqual(false, cdn.PrivateCdn.NoDefaultOrigins);
            Assert.AreEqual(OriginAction.Add, cdn.PrivateCdn.Origins[0].Action);
            Assert.AreEqual("sites/CDNPrivate/CDNFiles", cdn.PrivateCdn.Origins[0].Url);
        }

        [TestMethod]
        [TestCategory(TEST_CATEGORY)]
        public void XMLSerializer_Serialize_Tenant_ContentDeliveryNetwork()
        {
            var provider = new XMLFileSystemTemplateProvider($@"{AppDomain.CurrentDomain.BaseDirectory}\..\..\..\Resources", "Templates");

            var passWord = new SecureString();
            foreach (char c in "Pass@w0rd") passWord.AppendChar(c);

            var result = new ProvisioningTemplate { ParentHierarchy = new ProvisioningHierarchy() };

            var cdnSettings = new CdnSettings
            {
                Enabled = true,
                NoDefaultOrigins = true,
                ExcludeIfNoScriptDisabled = "JS,CSS",
                ExcludeRestrictedSiteClassifications = "HBI,GDPR",
                IncludeFileExtensions = "PDF,XML,JPG,JS,PNG",
            };
            cdnSettings.Origins.Add(new CdnOrigin
            {
                Action = OriginAction.Add,
                Url = "sites/CDN/CDNFiles"
            });
            result.ParentHierarchy.Tenant = new ProvisioningTenant(result.ApplicationLifecycleManagement.AppCatalog,
                new PnP.Framework.Provisioning.Model.ContentDeliveryNetwork(cdnSettings, cdnSettings));

            var serializer = new TargetSerializer();
            provider.SaveAs(result, TEST_OUT_FILE, serializer);

            var path = $"{provider.Connector.Parameters["ConnectionString"]}\\{provider.Connector.Parameters["Container"]}\\{TEST_OUT_FILE}";
            Assert.IsTrue(File.Exists(path)); path = Path.GetFullPath(path);
            var xml = XDocument.Load(path);
            var wrappedResult = XMLSerializer.Deserialize<TargetProvisioning>(xml);
            var cdn = wrappedResult.Tenant.ContentDeliveryNetwork;

            Assert.AreEqual(true, cdn.Public.Enabled);
            Assert.AreEqual("JS,CSS", cdn.Public.ExcludeIfNoScriptDisabled);
            Assert.AreEqual("HBI,GDPR", cdn.Public.ExcludeRestrictedSiteClassifications);
            Assert.AreEqual("PDF,XML,JPG,JS,PNG", cdn.Public.IncludeFileExtensions);
            Assert.AreEqual(true, cdn.Public.NoDefaultOrigins);
            Assert.AreEqual(CdnSettingOriginAction.Add, cdn.Public.Origins[0].Action);
            Assert.AreEqual("sites/CDN/CDNFiles", cdn.Public.Origins[0].Url);

            Assert.AreEqual(true, cdn.Private.Enabled);
            Assert.AreEqual("HBI,GDPR", cdn.Private.ExcludeRestrictedSiteClassifications);
            Assert.AreEqual("PDF,XML,JPG,JS,PNG", cdn.Private.IncludeFileExtensions);
            Assert.AreEqual(true, cdn.Private.NoDefaultOrigins);
            Assert.AreEqual(CdnSettingOriginAction.Add, cdn.Private.Origins[0].Action);
            Assert.AreEqual("sites/CDN/CDNFiles", cdn.Private.Origins[0].Url);
        }

        [TestMethod]
        [TestCategory(TEST_CATEGORY)]
        public void XMLSerializer_Deserialize_Tenant_SiteDesigns()
        {
            var provider = new XMLFileSystemTemplateProvider($@"{AppDomain.CurrentDomain.BaseDirectory}\..\..\..\Resources", "Templates");

            var hierarchy = provider.GetHierarchy(TEST_TEMPLATE);
            var siteDesigns = hierarchy.Tenant.SiteDesigns;

            Assert.AreEqual("Just a sample", siteDesigns[0].Description);
            Assert.AreEqual(true, siteDesigns[0].IsDefault);
            Assert.AreEqual(false, siteDesigns[0].Overwrite);
            Assert.AreEqual("PnP Site Design Preview", siteDesigns[0].PreviewImageAltText);
            Assert.AreEqual("PnPSiteDesign.png", siteDesigns[0].PreviewImageUrl);
            Assert.AreEqual("PnP Site Design", siteDesigns[0].Title);
            Assert.AreEqual(68, (int)siteDesigns[0].WebTemplate); // TenantHelper.ProcessSiteDesigns handles conversion to SiteDesignWebTemplate from valid integer

            Assert.AreEqual("user1@contoso.com", siteDesigns[0].Grants[0].Principal);
            Assert.AreEqual(SiteDesignRight.View, siteDesigns[0].Grants[0].Right);
            Assert.AreEqual(SiteDesignRight.None, siteDesigns[0].Grants[2].Right);

            Assert.AreEqual("{SiteScriptID:PnP Site Script 01}", siteDesigns[0].SiteScripts[0]);
        }

        [TestMethod]
        [TestCategory(TEST_CATEGORY)]
        public void XMLSerializer_Serialize_Tenant_SiteDesigns()
        {
            var provider = new XMLFileSystemTemplateProvider($@"{AppDomain.CurrentDomain.BaseDirectory}\..\..\..\Resources", "Templates");

            var passWord = new SecureString();
            foreach (char c in "Pass@w0rd") passWord.AppendChar(c);

            var result = new ProvisioningTemplate { ParentHierarchy = new ProvisioningHierarchy() };

            result.ParentHierarchy.Tenant = new ProvisioningTenant(result.ApplicationLifecycleManagement.AppCatalog,
                new PnP.Framework.Provisioning.Model.ContentDeliveryNetwork());

            result.Tenant.SiteDesigns.Add(new SiteDesign
            {
                Description = "Just a sample",
                IsDefault = true,
                Overwrite = false,
                PreviewImageAltText = "PnP Site Design Preview",
                PreviewImageUrl = "PnPSiteDesign.png",
                Title = "PnP Site Design",
                WebTemplate = SiteDesignWebTemplate.CommunicationSite,
                Grants =
                {
                    new SiteDesignGrant
                    {
                        Principal = "user1@contoso.com",
                        Right = SiteDesignRight.View
                    }
                },
                SiteScripts = new List<string> { "{SiteScriptID:PnP Site Script 01}" }
            });

            var serializer = new TargetSerializer();
            provider.SaveAs(result, TEST_OUT_FILE, serializer);

            var path = $"{provider.Connector.Parameters["ConnectionString"]}\\{provider.Connector.Parameters["Container"]}\\{TEST_OUT_FILE}";
            Assert.IsTrue(File.Exists(path)); path = Path.GetFullPath(path);
            var xml = XDocument.Load(path);
            var wrappedResult = XMLSerializer.Deserialize<TargetProvisioning>(xml);
            var siteDesigns = wrappedResult.Tenant.SiteDesigns;

            Assert.AreEqual("Just a sample", siteDesigns[0].Description);
            Assert.AreEqual(true, siteDesigns[0].IsDefault);
            Assert.AreEqual(false, siteDesigns[0].Overwrite);
            Assert.AreEqual("PnP Site Design Preview", siteDesigns[0].PreviewImageAltText);
            Assert.AreEqual("PnPSiteDesign.png", siteDesigns[0].PreviewImageUrl);
            Assert.AreEqual("PnP Site Design", siteDesigns[0].Title);
            Assert.AreEqual(SiteDesignsSiteDesignWebTemplate.CommunicationSite, siteDesigns[0].WebTemplate);

            Assert.AreEqual("user1@contoso.com", siteDesigns[0].Grants[0].Principal);
            Assert.AreEqual(SiteDesignsSiteDesignGrantRight.View, siteDesigns[0].Grants[0].Right);

            Assert.AreEqual("{SiteScriptID:PnP Site Script 01}", siteDesigns[0].SiteScripts[0].ID);
        }

        [TestMethod]
        [TestCategory(TEST_CATEGORY)]
        public void XMLSerializer_Deserialize_Tenant_SiteScripts()
        {
            var provider = new XMLFileSystemTemplateProvider($@"{AppDomain.CurrentDomain.BaseDirectory}\..\..\..\Resources", "Templates");
            var hierarchy = provider.GetHierarchy(TEST_TEMPLATE);
            var siteScripts = hierarchy.Tenant.SiteScripts;

            Assert.AreEqual("PnP Site Script Sample 01", siteScripts[0].Description);
            Assert.AreEqual(".\\pnp-site-script-01.json", siteScripts[0].JsonFilePath);
            Assert.AreEqual(true, siteScripts[0].Overwrite);
            Assert.AreEqual(false, siteScripts[1].Overwrite);
            Assert.AreEqual("PnP Site Script 01", siteScripts[0].Title);
        }

        [TestMethod]
        [TestCategory(TEST_CATEGORY)]
        public void XMLSerializer_Serialize_Tenant_SiteScripts()
        {
            var provider = new XMLFileSystemTemplateProvider($@"{AppDomain.CurrentDomain.BaseDirectory}\..\..\..\Resources", "Templates");

            var passWord = new SecureString();
            foreach (char c in "Pass@w0rd") passWord.AppendChar(c);

            var result = new ProvisioningTemplate { ParentHierarchy = new ProvisioningHierarchy() };

            result.ParentHierarchy.Tenant = new ProvisioningTenant(result.ApplicationLifecycleManagement.AppCatalog,
                new PnP.Framework.Provisioning.Model.ContentDeliveryNetwork());

            result.Tenant.SiteScripts.Add(new SiteScript
            {
                Description = "PnP Site Script Sample 01",
                Overwrite = true,
                Title = "PnP Site Script 01",
                JsonFilePath = ".\\pnp-site-script-01.json"
            });

            result.Tenant.SiteScripts.Add(new SiteScript
            {
                Description = "PnP Site Script Sample 02",
                Overwrite = false,
                Title = "PnP Site Script 02",
                JsonFilePath = ".\\pnp-site-script-02.json"
            });

            var serializer = new TargetSerializer();
            provider.SaveAs(result, TEST_OUT_FILE, serializer);

            var path = $"{provider.Connector.Parameters["ConnectionString"]}\\{provider.Connector.Parameters["Container"]}\\{TEST_OUT_FILE}";
            Assert.IsTrue(File.Exists(path)); path = Path.GetFullPath(path);
            var xml = XDocument.Load(path);
            var wrappedResult = XMLSerializer.Deserialize<TargetProvisioning>(xml);
            var siteScripts = wrappedResult.Tenant.SiteScripts;

            Assert.AreEqual("PnP Site Script Sample 01", siteScripts[0].Description);
            Assert.AreEqual(".\\pnp-site-script-01.json", siteScripts[0].JsonFilePath);
            Assert.AreEqual(true, siteScripts[0].Overwrite);
            Assert.AreEqual(false, siteScripts[1].Overwrite);
            Assert.AreEqual("PnP Site Script 01", siteScripts[0].Title);
        }

        [TestMethod]
        [TestCategory(TEST_CATEGORY)]
        public void XMLSerializer_Deserialize_Tenant_StorageEntities()
        {
            var provider = new XMLFileSystemTemplateProvider($@"{AppDomain.CurrentDomain.BaseDirectory}\..\..\..\Resources", "Templates");

            var hierarchy = provider.GetHierarchy(TEST_TEMPLATE);
            var storageEntities = hierarchy.Tenant.StorageEntities;

            Assert.AreEqual("Description 01", storageEntities[0].Description);
            Assert.AreEqual("Comment 01", storageEntities[0].Comment);
            Assert.AreEqual("PnPKey01", storageEntities[0].Key);
            Assert.AreEqual("My custom tenant-wide value 01", storageEntities[0].Value);
        }

        [TestMethod]
        [TestCategory(TEST_CATEGORY)]
        public void XMLSerializer_Serialize_Tenant_StorageEntities()
        {
            var provider = new XMLFileSystemTemplateProvider($@"{AppDomain.CurrentDomain.BaseDirectory}\..\..\..\Resources", "Templates");

            var passWord = new SecureString();
            foreach (char c in "Pass@w0rd") passWord.AppendChar(c);

            var result = new ProvisioningTemplate { ParentHierarchy = new ProvisioningHierarchy() };

            result.ParentHierarchy.Tenant = new ProvisioningTenant(result.ApplicationLifecycleManagement.AppCatalog,
                new PnP.Framework.Provisioning.Model.ContentDeliveryNetwork());

            result.Tenant.StorageEntities.Add(new StorageEntity
            {
                Description = "Description 01",
                Key = "PnPKey01",
                Value = "My custom tenant-wide value 01",
                Comment = "Comment 01"
            });

            var serializer = new TargetSerializer();
            provider.SaveAs(result, TEST_OUT_FILE, serializer);

            var path = $"{provider.Connector.Parameters["ConnectionString"]}\\{provider.Connector.Parameters["Container"]}\\{TEST_OUT_FILE}";
            Assert.IsTrue(File.Exists(path)); path = Path.GetFullPath(path);
            var xml = XDocument.Load(path);
            var wrappedResult = XMLSerializer.Deserialize<TargetProvisioning>(xml);
            var storageEntities = wrappedResult.Tenant.StorageEntities;

            Assert.AreEqual("Description 01", storageEntities[0].Description);
            Assert.AreEqual("Comment 01", storageEntities[0].Comment);
            Assert.AreEqual("PnPKey01", storageEntities[0].Key);
            Assert.AreEqual("My custom tenant-wide value 01", storageEntities[0].Value);
        }

        [TestMethod]
        [TestCategory(TEST_CATEGORY)]
        public void XMLSerializer_Deserialize_Tenant_Themes()
        {
            var provider = new XMLFileSystemTemplateProvider($@"{AppDomain.CurrentDomain.BaseDirectory}\..\..\..\Resources", "Templates");

            var hierarchy = provider.GetHierarchy(TEST_TEMPLATE);
            var themes = hierarchy.Tenant.Themes;

            Assert.AreEqual(false, themes[0].IsInverted);
            Assert.AreEqual(false, themes[0].Overwrite);
            Assert.AreEqual("CustomOrange", themes[0].Name);
            Assert.IsTrue(themes[0].Palette.Contains("\"neutralQuaternaryAlt\": \"#dadada\""));

        }

        [TestMethod]
        [TestCategory(TEST_CATEGORY)]
        public void XMLSerializer_Serialize_Tenant_Themes()
        {
            var provider = new XMLFileSystemTemplateProvider($@"{AppDomain.CurrentDomain.BaseDirectory}\..\..\..\Resources", "Templates");

            var passWord = new SecureString();
            foreach (char c in "Pass@w0rd") passWord.AppendChar(c);

            var result = new ProvisioningTemplate { ParentHierarchy = new ProvisioningHierarchy() };

            result.ParentHierarchy.Tenant = new ProvisioningTenant(result.ApplicationLifecycleManagement.AppCatalog,
                new PnP.Framework.Provisioning.Model.ContentDeliveryNetwork());

            result.Tenant.Themes.Add(new PnP.Framework.Provisioning.Model.Theme
            {
                Name = "CustomOrange",
                IsInverted = false,
                Palette = "{\"neutralQuaternaryAlt\": \"#dadada\"}",
                Overwrite = false
            });

            var serializer = new TargetSerializer();
            provider.SaveAs(result, TEST_OUT_FILE, serializer);

            var path = $"{provider.Connector.Parameters["ConnectionString"]}\\{provider.Connector.Parameters["Container"]}\\{TEST_OUT_FILE}";
            Assert.IsTrue(File.Exists(path)); path = Path.GetFullPath(path);
            var xml = XDocument.Load(path);
            var wrappedResult = XMLSerializer.Deserialize<TargetProvisioning>(xml);
            var themes = wrappedResult.Tenant.Themes;

            Assert.AreEqual(false, themes[0].IsInverted);
            Assert.AreEqual(false, themes[0].Overwrite);
            Assert.AreEqual("CustomOrange", themes[0].Name);
            Assert.IsTrue(themes[0].Text[0].Contains("\"neutralQuaternaryAlt\": \"#dadada\""));
        }

        [TestMethod]
        [TestCategory(TEST_CATEGORY)]
        public void XMLSerializer_Deserialize_Tenant_SPUserProfiles()
        {
            var provider = new XMLFileSystemTemplateProvider($@"{AppDomain.CurrentDomain.BaseDirectory}\..\..\..\Resources", "Templates");

            var hierarchy = provider.GetHierarchy(TEST_TEMPLATE);
            var ups = hierarchy.Tenant.SPUsersProfiles;

            Assert.AreEqual(4, ups.Count);
            Assert.AreEqual("user1@contoso.com", ups[0].TargetUser);
            Assert.AreEqual("IT", ups[0].Properties["Department"]);
            Assert.AreEqual("Milan, Italy", ups[0].Properties["OfficeLocation"]);
            Assert.AreEqual("+39-030-123456", ups[0].Properties["PhoneNumber"]);

            Assert.AreEqual("user2@{parameter:O365TenantName}.onmicrosoft.com", ups[1].TargetUser);
            Assert.AreEqual("HR", ups[1].Properties["Department"]);
            Assert.AreEqual("Seattle, WS", ups[1].Properties["OfficeLocation"]);
            Assert.AreEqual("+1-321-123456", ups[1].Properties["PhoneNumber"]);

            Assert.AreEqual("group01@contoso.com", ups[2].TargetGroup);
            Assert.AreEqual("R&D", ups[2].Properties["Department"]);
            Assert.AreEqual("New York, NY", ups[2].Properties["OfficeLocation"]);
            Assert.AreEqual("+1-456-123456", ups[2].Properties["PhoneNumber"]);

            Assert.AreEqual("group02@{parameter:O365TenantName}.onmicrosoft.com", ups[3].TargetGroup);
            Assert.AreEqual("Production", ups[3].Properties["Department"]);
            Assert.AreEqual("Washington, DC", ups[3].Properties["OfficeLocation"]);
            Assert.AreEqual("+1-789-123456", ups[3].Properties["PhoneNumber"]);
        }

        [TestMethod]
        [TestCategory(TEST_CATEGORY)]
        public void XMLSerializer_Serialize_Tenant_SPUserProfiles()
        {
            var provider = new XMLFileSystemTemplateProvider($@"{AppDomain.CurrentDomain.BaseDirectory}\..\..\..\Resources", "Templates");

            var result = new ProvisioningTemplate { ParentHierarchy = new ProvisioningHierarchy() };

            result.ParentHierarchy.Tenant = new ProvisioningTenant(result.ApplicationLifecycleManagement.AppCatalog,
                new PnP.Framework.Provisioning.Model.ContentDeliveryNetwork());

            result.Tenant.SPUsersProfiles.AddRange(new PnP.Framework.Provisioning.Model.SPUPS.UserProfile[]
            {
                new PnP.Framework.Provisioning.Model.SPUPS.UserProfile
                {
                    TargetUser = "user1@contoso.com",
                },
                new PnP.Framework.Provisioning.Model.SPUPS.UserProfile
                {
                    TargetGroup = "group01@contoso.com",
                },
            });
            result.Tenant.SPUsersProfiles[0].Properties.Add("Department", "IT");
            result.Tenant.SPUsersProfiles[0].Properties.Add("OfficeLocation", "Milan, Italy");
            result.Tenant.SPUsersProfiles[0].Properties.Add("PhoneNumber", "+39-030-123456");
            result.Tenant.SPUsersProfiles[1].Properties.Add("Department", "R&D");
            result.Tenant.SPUsersProfiles[1].Properties.Add("OfficeLocation", "New York, NY");
            result.Tenant.SPUsersProfiles[1].Properties.Add("PhoneNumber", "+1-456-123456");

            var serializer = new TargetSerializer();
            provider.SaveAs(result, TEST_OUT_FILE, serializer);

            var path = $"{provider.Connector.Parameters["ConnectionString"]}\\{provider.Connector.Parameters["Container"]}\\{TEST_OUT_FILE}";
            Assert.IsTrue(File.Exists(path)); path = Path.GetFullPath(path);
            var xml = XDocument.Load(path);
            var wrappedResult = XMLSerializer.Deserialize<TargetProvisioning>(xml);
            var ups = wrappedResult.Tenant.SPUsersProfiles;

            Assert.AreEqual(2, ups.Length);
            Assert.AreEqual("user1@contoso.com", ups[0].TargetUser);
            Assert.AreEqual("Department", ups[0].Property[0].Key);
            Assert.AreEqual("IT", ups[0].Property[0].Value);
            Assert.AreEqual("OfficeLocation", ups[0].Property[1].Key);
            Assert.AreEqual("Milan, Italy", ups[0].Property[1].Value);
            Assert.AreEqual("PhoneNumber", ups[0].Property[2].Key);
            Assert.AreEqual("+39-030-123456", ups[0].Property[2].Value);
            Assert.AreEqual("group01@contoso.com", ups[1].TargetGroup);
            Assert.AreEqual("Department", ups[1].Property[0].Key);
            Assert.AreEqual("R&D", ups[1].Property[0].Value);
            Assert.AreEqual("OfficeLocation", ups[1].Property[1].Key);
            Assert.AreEqual("New York, NY", ups[1].Property[1].Value);
            Assert.AreEqual("PhoneNumber", ups[1].Property[2].Key);
            Assert.AreEqual("+1-456-123456", ups[1].Property[2].Value);
        }

        [TestMethod]
        [TestCategory(TEST_CATEGORY)]
        public void XMLSerializer_Deserialize_Tenant_Office365GroupsLifecyclePolicies()
        {
            var provider = new XMLFileSystemTemplateProvider($@"{AppDomain.CurrentDomain.BaseDirectory}\..\..\..\Resources", "Templates");

            var hierarchy = provider.GetHierarchy(TEST_TEMPLATE);
            var o365LifecyclePolicies = hierarchy.Tenant.Office365GroupLifecyclePolicies;

            Assert.AreEqual(1, o365LifecyclePolicies.Count);
            Assert.AreEqual("GROUP_LIFECYCLE_01", o365LifecyclePolicies[0].ID);
            Assert.AreEqual(180, o365LifecyclePolicies[0].GroupLifetimeInDays);
            Assert.AreEqual("admin01@contoso.com,admin02@{parameter:O365TenantName}.onmicrosoft.com", o365LifecyclePolicies[0].AlternateNotificationEmails);
            Assert.AreEqual(PnP.Framework.Provisioning.Model.Office365Groups.ManagedGroupTypes.Selected,
                o365LifecyclePolicies[0].ManagedGroupTypes);
        }

        [TestMethod]
        [TestCategory(TEST_CATEGORY)]
        public void XMLSerializer_Serialize_Tenant_Office365GroupsLifecyclePolicies()
        {
            var provider = new XMLFileSystemTemplateProvider($@"{AppDomain.CurrentDomain.BaseDirectory}\..\..\..\Resources", "Templates");

            var result = new ProvisioningTemplate { ParentHierarchy = new ProvisioningHierarchy() };

            result.ParentHierarchy.Tenant = new ProvisioningTenant(result.ApplicationLifecycleManagement.AppCatalog,
                new PnP.Framework.Provisioning.Model.ContentDeliveryNetwork());

            result.Tenant.Office365GroupLifecyclePolicies.Add(
                new PnP.Framework.Provisioning.Model.Office365Groups.Office365GroupLifecyclePolicy
                {
                    ID = "GROUP_LIFECYCLE_01",
                    GroupLifetimeInDays = 180,
                    AlternateNotificationEmails = "admin01@contoso.com,admin02@{parameter:O365TenantName}.onmicrosoft.com",
                    ManagedGroupTypes = PnP.Framework.Provisioning.Model.Office365Groups.ManagedGroupTypes.Selected,
                }
                );

            var serializer = new TargetSerializer();
            provider.SaveAs(result, TEST_OUT_FILE, serializer);

            var path = $"{provider.Connector.Parameters["ConnectionString"]}\\{provider.Connector.Parameters["Container"]}\\{TEST_OUT_FILE}";
            Assert.IsTrue(File.Exists(path)); path = Path.GetFullPath(path);
            var xml = XDocument.Load(path);
            var wrappedResult = XMLSerializer.Deserialize<TargetProvisioning>(xml);
            var o365LifecyclePolicies = wrappedResult.Tenant.Office365GroupLifecyclePolicies;

            Assert.AreEqual(1, o365LifecyclePolicies.Length);
            Assert.AreEqual("GROUP_LIFECYCLE_01", o365LifecyclePolicies[0].ID);
            Assert.AreEqual(180, o365LifecyclePolicies[0].GroupLifetimeInDays);
            Assert.AreEqual("admin01@contoso.com,admin02@{parameter:O365TenantName}.onmicrosoft.com",
                o365LifecyclePolicies[0].AlternateNotificationEmails);
            Assert.AreEqual(Office365GroupLifecyclePolicyManagedGroupTypes.Selected,
                o365LifecyclePolicies[0].ManagedGroupTypes);
        }

        [TestMethod]
        [TestCategory(TEST_CATEGORY)]
        public void XMLSerializer_Deserialize_Tenant_Office365GroupsSettings()
        {
            var provider = new XMLFileSystemTemplateProvider($@"{AppDomain.CurrentDomain.BaseDirectory}\..\..\..\Resources", "Templates");

            var hierarchy = provider.GetHierarchy(TEST_TEMPLATE);
            var o365GroupsSettings = hierarchy.Tenant.Office365GroupsSettings;

            Assert.AreEqual(4, o365GroupsSettings.Properties.Count);
            Assert.AreEqual("http://aka.ms/SharePointPnP", o365GroupsSettings.Properties["UsageGuidelinesUrl"]);
            Assert.AreEqual("HBI,MBI,LBI,GDPR,TopSecret", o365GroupsSettings.Properties["ClassificationList"]);
            Assert.AreEqual("true", o365GroupsSettings.Properties["EnableGroupCreation"]);
            Assert.AreEqual("MBI", o365GroupsSettings.Properties["DefaultClassification"]);
        }

        [TestMethod]
        [TestCategory(TEST_CATEGORY)]
        public void XMLSerializer_Serialize_Tenant_Office365GroupsSettings()
        {
            var provider = new XMLFileSystemTemplateProvider($@"{AppDomain.CurrentDomain.BaseDirectory}\..\..\..\Resources", "Templates");

            var result = new ProvisioningTemplate { ParentHierarchy = new ProvisioningHierarchy() };

            result.ParentHierarchy.Tenant = new ProvisioningTenant(result.ApplicationLifecycleManagement.AppCatalog,
                new PnP.Framework.Provisioning.Model.ContentDeliveryNetwork());

            result.Tenant.Office365GroupsSettings.Properties.Add("UsageGuidelinesUrl", "http://aka.ms/SharePointPnP");
            result.Tenant.Office365GroupsSettings.Properties.Add("ClassificationList", "HBI,MBI,LBI,GDPR,TopSecret");
            result.Tenant.Office365GroupsSettings.Properties.Add("EnableGroupCreation", "true");
            result.Tenant.Office365GroupsSettings.Properties.Add("DefaultClassification", "MBI");

            var serializer = new TargetSerializer();
            provider.SaveAs(result, TEST_OUT_FILE, serializer);

            var path = $"{provider.Connector.Parameters["ConnectionString"]}\\{provider.Connector.Parameters["Container"]}\\{TEST_OUT_FILE}";
            Assert.IsTrue(File.Exists(path)); path = Path.GetFullPath(path);
            var xml = XDocument.Load(path);
            var wrappedResult = XMLSerializer.Deserialize<TargetProvisioning>(xml);
            var o365GroupsSettings = wrappedResult.Tenant.Office365GroupsSettings;

            Assert.AreEqual(4, o365GroupsSettings.Length);
            Assert.AreEqual("UsageGuidelinesUrl", o365GroupsSettings[0].Key);
            Assert.AreEqual("http://aka.ms/SharePointPnP", o365GroupsSettings[0].Value);
            Assert.AreEqual("ClassificationList", o365GroupsSettings[1].Key);
            Assert.AreEqual("HBI,MBI,LBI,GDPR,TopSecret", o365GroupsSettings[1].Value);
            Assert.AreEqual("EnableGroupCreation", o365GroupsSettings[2].Key);
            Assert.AreEqual("true", o365GroupsSettings[2].Value);
            Assert.AreEqual("DefaultClassification", o365GroupsSettings[3].Key);
            Assert.AreEqual("MBI", o365GroupsSettings[3].Value);
        }

        [TestMethod]
        [TestCategory(TEST_CATEGORY)]
        public void XMLSerializer_Deserialize_Tenant_SharingSettings()
        {
            var provider = new XMLFileSystemTemplateProvider($@"{AppDomain.CurrentDomain.BaseDirectory}\..\..\..\Resources", "Templates");

            var hierarchy = provider.GetHierarchy(TEST_TEMPLATE);
            var sharingSettings = hierarchy.Tenant.SharingSettings;

            Assert.AreEqual(SharingCapability.ExternalUserAndGuestSharing, sharingSettings.SharingCapability);
            Assert.AreEqual(30, sharingSettings.RequireAnonymousLinksExpireInDays);
            Assert.AreEqual(PnP.Framework.Provisioning.Model.AnonymousLinkType.View, sharingSettings.FileAnonymousLinkType);
            Assert.AreEqual(PnP.Framework.Provisioning.Model.AnonymousLinkType.Edit, sharingSettings.FolderAnonymousLinkType);
            Assert.AreEqual(SharingLinkType.AnonymousAccess, sharingSettings.DefaultSharingLinkType);
            Assert.AreEqual(true, sharingSettings.PreventExternalUsersFromResharing);
            Assert.AreEqual(true, sharingSettings.RequireAcceptingAccountMatchInvitedAccount);
            Assert.AreEqual(PnP.Framework.Provisioning.Model.SharingDomainRestrictionMode.AllowList, sharingSettings.SharingDomainRestrictionMode);
            Assert.AreEqual(2, sharingSettings.AllowedDomainList.Count);
            Assert.AreEqual("contoso.com", sharingSettings.AllowedDomainList[0]);
            Assert.AreEqual("contoso-elettronics.com", sharingSettings.AllowedDomainList[1]);
            Assert.AreEqual(2, sharingSettings.BlockedDomainList.Count);
            Assert.AreEqual("contoso.com", sharingSettings.BlockedDomainList[0]);
            Assert.AreEqual("contoso-elettronics.com", sharingSettings.BlockedDomainList[1]);
        }

        [TestMethod]
        [TestCategory(TEST_CATEGORY)]
        public void XMLSerializer_Serialize_Tenant_SharingSettings()
        {
            var provider = new XMLFileSystemTemplateProvider($@"{AppDomain.CurrentDomain.BaseDirectory}\..\..\..\Resources", "Templates");

            var result = new ProvisioningTemplate { ParentHierarchy = new ProvisioningHierarchy() };

            result.ParentHierarchy.Tenant = new ProvisioningTenant(result.ApplicationLifecycleManagement.AppCatalog,
                new PnP.Framework.Provisioning.Model.ContentDeliveryNetwork());

            result.Tenant.SharingSettings = new PnP.Framework.Provisioning.Model.SharingSettings
            {
                SharingCapability = SharingCapability.ExternalUserAndGuestSharing,
                RequireAnonymousLinksExpireInDays = 30,
                FileAnonymousLinkType = PnP.Framework.Provisioning.Model.AnonymousLinkType.View,
                FolderAnonymousLinkType = PnP.Framework.Provisioning.Model.AnonymousLinkType.Edit,
                DefaultSharingLinkType = SharingLinkType.AnonymousAccess,
                PreventExternalUsersFromResharing = true,
                RequireAcceptingAccountMatchInvitedAccount = true,
                SharingDomainRestrictionMode = SharingDomainRestrictionMode.AllowList
            };
            result.Tenant.SharingSettings.AllowedDomainList.Add("contoso.com");
            result.Tenant.SharingSettings.AllowedDomainList.Add("contoso-elettronics.com");
            result.Tenant.SharingSettings.BlockedDomainList.Add("contoso.com");
            result.Tenant.SharingSettings.BlockedDomainList.Add("contoso-elettronics.com");

            var serializer = new TargetSerializer();
            provider.SaveAs(result, TEST_OUT_FILE, serializer);

            var path = $"{provider.Connector.Parameters["ConnectionString"]}\\{provider.Connector.Parameters["Container"]}\\{TEST_OUT_FILE}";
            Assert.IsTrue(File.Exists(path)); path = Path.GetFullPath(path);
            var xml = XDocument.Load(path);
            var wrappedResult = XMLSerializer.Deserialize<TargetProvisioning>(xml);
            var sharingSettings = wrappedResult.Tenant.SharingSettings;

            Assert.AreEqual(TargetNamespace.SharingSettingsSharingCapability.ExternalUserAndGuestSharing, sharingSettings.SharingCapability);
            Assert.AreEqual(30, sharingSettings.RequireAnonymousLinksExpireInDays);
            Assert.AreEqual(TargetNamespace.AnonymousLinkType.View, sharingSettings.FileAnonymousLinkType);
            Assert.AreEqual(TargetNamespace.AnonymousLinkType.Edit, sharingSettings.FolderAnonymousLinkType);
            Assert.AreEqual(TargetNamespace.SharingSettingsDefaultSharingLinkType.AnonymousAccess, sharingSettings.DefaultSharingLinkType);
            Assert.AreEqual(true, sharingSettings.PreventExternalUsersFromResharing);
            Assert.AreEqual(true, sharingSettings.RequireAcceptingAccountMatchInvitedAccount);
            Assert.AreEqual(TargetNamespace.SharingSettingsSharingDomainRestrictionMode.AllowList, sharingSettings.SharingDomainRestrictionMode);
            Assert.AreEqual("contoso.com,contoso-elettronics.com", sharingSettings.AllowedDomainList);
            Assert.AreEqual("contoso.com,contoso-elettronics.com", sharingSettings.BlockedDomainList);
        }

        [TestMethod]
        [TestCategory(TEST_CATEGORY)]
        public void XMLSerializer_Deserialize_Localizations()
        {
            var provider = new XMLFileSystemTemplateProvider($@"{AppDomain.CurrentDomain.BaseDirectory}\..\..\..\Resources", "Templates");

            var serializer = new TargetSerializer();
            var template = provider.GetTemplate(TEST_TEMPLATE, serializer);
            var localizations = template.Localizations;

            Assert.IsNotNull(localizations);
            Assert.AreEqual(1040, localizations.DefaultLCID);
            Assert.AreEqual(6, localizations.Count);
            Assert.AreEqual(1033, localizations[0].LCID);
            Assert.AreEqual("core", localizations[0].Name);
            Assert.AreEqual("resources-core.en-us.resx", localizations[0].ResourceFile);
            Assert.AreEqual(1040, localizations[1].LCID);
            Assert.AreEqual("core", localizations[1].Name);
            Assert.AreEqual("resources-core.it-it.resx", localizations[1].ResourceFile);
            Assert.AreEqual(1031, localizations[2].LCID);
            Assert.AreEqual("core", localizations[2].Name);
            Assert.AreEqual("resources-core.de-de.resx", localizations[2].ResourceFile);
            Assert.AreEqual(1033, localizations[3].LCID);
            Assert.AreEqual("intranet", localizations[3].Name);
            Assert.AreEqual("resources-intranet.en-us.resx", localizations[3].ResourceFile);
            Assert.AreEqual(1040, localizations[4].LCID);
            Assert.AreEqual("intranet", localizations[4].Name);
            Assert.AreEqual("resources-intranet.it-it.resx", localizations[4].ResourceFile);
            Assert.AreEqual(1031, localizations[5].LCID);
            Assert.AreEqual("intranet", localizations[5].Name);
            Assert.AreEqual("resources-intranet.de-de.resx", localizations[5].ResourceFile);
        }

        [TestMethod]
        [TestCategory(TEST_CATEGORY)]
        public void XMLSerializer_Serialize_Localizations()
        {
            var provider = new XMLFileSystemTemplateProvider($@"{AppDomain.CurrentDomain.BaseDirectory}\..\..\..\Resources", "Templates");

            var result = new ProvisioningTemplate();

            result.Localizations.DefaultLCID = 1040;
            result.Localizations.Add(new Localization { LCID = 1040, Name = "core", ResourceFile = "resources-core.it-it.resx" });

            var serializer = new TargetSerializer();
            provider.SaveAs(result, TEST_OUT_FILE, serializer);

            var path = $"{provider.Connector.Parameters["ConnectionString"]}\\{provider.Connector.Parameters["Container"]}\\{TEST_OUT_FILE}";
            Assert.IsTrue(File.Exists(path)); path = Path.GetFullPath(path);
            var xml = XDocument.Load(path);
            var wrappedResult = XMLSerializer.Deserialize<TargetProvisioning>(xml);

            var localizations = wrappedResult.Localizations;

            Assert.IsNotNull(localizations);
            Assert.AreEqual(1040, localizations.DefaultLCID);
            Assert.AreEqual(1, localizations.Localization.Length);
            Assert.AreEqual(1040, localizations.Localization[0].LCID);
            Assert.AreEqual("core", localizations.Localization[0].Name);
            Assert.AreEqual("resources-core.it-it.resx", localizations.Localization[0].ResourceFile);
        }

        [TestMethod]
        [TestCategory(TEST_CATEGORY)]
        public void XMLSerializer_Deserialize_Properties()
        {
            var provider = new XMLFileSystemTemplateProvider($@"{AppDomain.CurrentDomain.BaseDirectory}\..\..\..\Resources", "Templates");

            var serializer = new TargetSerializer();
            var template = provider.GetTemplate(TEST_TEMPLATE, serializer);
            var properties = template.Properties;

            Assert.IsTrue(properties.ContainsKey("Something"));
            Assert.AreEqual("One property", properties["Something"]);
        }

        [TestMethod]
        [TestCategory(TEST_CATEGORY)]
        public void XMLSerializer_Serialize_Properties()
        {
            var provider = new XMLFileSystemTemplateProvider($@"{AppDomain.CurrentDomain.BaseDirectory}\..\..\..\Resources", "Templates");

            var result = new ProvisioningTemplate();

            result.Properties.Add("Something", "One property");


            var serializer = new TargetSerializer();
            provider.SaveAs(result, TEST_OUT_FILE, serializer);

            var path = $"{provider.Connector.Parameters["ConnectionString"]}\\{provider.Connector.Parameters["Container"]}\\{TEST_OUT_FILE}";
            Assert.IsTrue(File.Exists(path)); path = Path.GetFullPath(path);
            var xml = XDocument.Load(path);
            var wrappedResult = XMLSerializer.Deserialize<TargetProvisioning>(xml);

            var template = wrappedResult.Templates[0].ProvisioningTemplate.First();

            var properties = template.Properties;

            Assert.IsTrue(properties.Where(p => p.Key.Equals("Something")).Count() == 1);
            Assert.AreEqual("One property", properties.Single(p => p.Key.Equals("Something")).Value);
        }

        [TestMethod]
        [TestCategory(TEST_CATEGORY)]
        public void XMLSerializer_Deserialize_SitePolicy()
        {
            var provider = new XMLFileSystemTemplateProvider($@"{AppDomain.CurrentDomain.BaseDirectory}\..\..\..\Resources", "Templates");

            var serializer = new TargetSerializer();
            var template = provider.GetTemplate(TEST_TEMPLATE, serializer);
            var sitePolicy = template.SitePolicy;

            Assert.AreEqual("HBI", sitePolicy);
        }

        [TestMethod]
        [TestCategory(TEST_CATEGORY)]
        public void XMLSerializer_Serialize_SitePolicy()
        {
            var provider = new XMLFileSystemTemplateProvider($@"{AppDomain.CurrentDomain.BaseDirectory}\..\..\..\Resources", "Templates");

            var result = new ProvisioningTemplate { SitePolicy = "HBI" };

            var serializer = new TargetSerializer();
            provider.SaveAs(result, TEST_OUT_FILE, serializer);

            var path = $"{provider.Connector.Parameters["ConnectionString"]}\\{provider.Connector.Parameters["Container"]}\\{TEST_OUT_FILE}";
            Assert.IsTrue(File.Exists(path)); path = Path.GetFullPath(path);
            var xml = XDocument.Load(path);
            var wrappedResult = XMLSerializer.Deserialize<TargetProvisioning>(xml);

            var template = wrappedResult.Templates[0].ProvisioningTemplate.First();

            var sitePolicy = template.SitePolicy;

            Assert.AreEqual("HBI", sitePolicy);
        }

        [TestMethod]
        [TestCategory(TEST_CATEGORY)]
        public void XMLSerializer_Deserialize_WebSettings()
        {
            var provider = new XMLFileSystemTemplateProvider($@"{AppDomain.CurrentDomain.BaseDirectory}\..\..\..\Resources", "Templates");

            var serializer = new TargetSerializer();
            var template = provider.GetTemplate(TEST_TEMPLATE, serializer);
            var webSettings = template.WebSettings;

            Assert.AreEqual("someone@company.com", webSettings.RequestAccessEmail);
            Assert.AreEqual(false, webSettings.NoCrawl);
            Assert.AreEqual("/Pages/Home.aspx", webSettings.WelcomePage);
            Assert.AreEqual("Site Title", webSettings.Title);
            Assert.AreEqual("Site Description", webSettings.Description);
            Assert.AreEqual("{sitecollection}/SiteAssets/Logo.png", webSettings.SiteLogo);
            Assert.AreEqual("{sitecollection}/SiteAssets/Logo.png", webSettings.SiteLogoThumbnail);
            Assert.AreEqual("{sitecollection}/Resources/Themes/Contoso/Contoso.css", webSettings.AlternateCSS);
            Assert.AreEqual("{sitecollection}/_catalogs/MasterPage/oslo.master", webSettings.MasterPageUrl);
            Assert.AreEqual("{sitecollection}/_catalogs/MasterPage/CustomMaster.master", webSettings.CustomMasterPageUrl);
            Assert.AreEqual("/sites/hubsite", webSettings.HubSiteUrl);
            Assert.AreEqual(false, webSettings.CommentsOnSitePagesDisabled);
            Assert.AreEqual(true, webSettings.QuickLaunchEnabled);
            Assert.AreEqual(true, webSettings.IsMultilingual);
            Assert.AreEqual(true, webSettings.OverwriteTranslationsOnChange);
            Assert.IsNotNull(webSettings.AlternateUICultures);
            Assert.AreEqual(3, webSettings.AlternateUICultures.Count);
            Assert.AreEqual(1033, webSettings.AlternateUICultures[0].LCID);
            Assert.AreEqual(1040, webSettings.AlternateUICultures[1].LCID);
            Assert.AreEqual(1035, webSettings.AlternateUICultures[2].LCID);
            Assert.AreEqual(true, webSettings.ExcludeFromOfflineClient);
            Assert.AreEqual(true, webSettings.MembersCanShare);
            Assert.AreEqual(false, webSettings.DisableFlows);
            Assert.AreEqual(false, webSettings.DisableAppViews);
            Assert.AreEqual(true, webSettings.HorizontalQuickLaunch);
            Assert.AreEqual(SearchScopes.Hub, webSettings.SearchScope);
            Assert.AreEqual(PnP.Framework.Provisioning.Model.SearchBoxInNavBar.AllPages, webSettings.SearchBoxInNavBar);
            Assert.AreEqual("https://contoso.sharepoint.com/sites/SearchCenter", webSettings.SearchCenterUrl);
        }

        [TestMethod]
        [TestCategory(TEST_CATEGORY)]
        public void XMLSerializer_Serialize_WebSettings()
        {
            var provider = new XMLFileSystemTemplateProvider($@"{AppDomain.CurrentDomain.BaseDirectory}\..\..\..\Resources", "Templates");

            var result = new ProvisioningTemplate
            {
                WebSettings = new PnP.Framework.Provisioning.Model.WebSettings
                {
                    RequestAccessEmail = "someone@company.com",
                    NoCrawl = false,
                    WelcomePage = "/Pages/Home.aspx",
                    Title = "Site Title",
                    Description = "Site Description",
                    SiteLogo = "{sitecollection}/SiteAssets/Logo.png",
                    SiteLogoThumbnail = "{sitecollection}/SiteAssets/Logo.png",
                    AlternateCSS = "{sitecollection}/Resources/Themes/Contoso/Contoso.css",
                    MasterPageUrl = "{sitecollection}/_catalogs/MasterPage/oslo.master",
                    CustomMasterPageUrl = "{sitecollection}/_catalogs/MasterPage/CustomMaster.master",
                    HubSiteUrl = "/sites/hubsite",
                    CommentsOnSitePagesDisabled = false,
                    QuickLaunchEnabled = true,
                    IsMultilingual = true,
                    OverwriteTranslationsOnChange = true,
                    ExcludeFromOfflineClient = true,
                    MembersCanShare = true,
                    DisableFlows = false,
                    DisableAppViews = false,
                    HorizontalQuickLaunch = true,
                    SearchScope = SearchScopes.Hub,
                    SearchBoxInNavBar = PnP.Framework.Provisioning.Model.SearchBoxInNavBar.AllPages,
                    SearchCenterUrl = "https://contoso.sharepoint.com/sites/SearchCenter",
                }
            };

            result.WebSettings.AlternateUICultures.AddRange(
                new AlternateUICulture[]
                {
                    new AlternateUICulture { LCID = 1033 },
                    new AlternateUICulture { LCID = 1040 },
                    new AlternateUICulture { LCID = 1035 },
                });

            var serializer = new TargetSerializer();
            provider.SaveAs(result, TEST_OUT_FILE, serializer);

            var path = $"{provider.Connector.Parameters["ConnectionString"]}\\{provider.Connector.Parameters["Container"]}\\{TEST_OUT_FILE}";
            Assert.IsTrue(File.Exists(path)); path = Path.GetFullPath(path);
            var xml = XDocument.Load(path);
            var wrappedResult = XMLSerializer.Deserialize<TargetProvisioning>(xml);

            var template = wrappedResult.Templates[0].ProvisioningTemplate.First();

            var webSettings = template.WebSettings;

            Assert.AreEqual("someone@company.com", webSettings.RequestAccessEmail);
            Assert.AreEqual(false, webSettings.NoCrawl);
            Assert.AreEqual("/Pages/Home.aspx", webSettings.WelcomePage);
            Assert.AreEqual("Site Title", webSettings.Title);
            Assert.AreEqual("Site Description", webSettings.Description);
            Assert.AreEqual("{sitecollection}/SiteAssets/Logo.png", webSettings.SiteLogo);
            Assert.AreEqual("{sitecollection}/SiteAssets/Logo.png", webSettings.SiteLogoThumbnail);
            Assert.AreEqual("{sitecollection}/Resources/Themes/Contoso/Contoso.css", webSettings.AlternateCSS);
            Assert.AreEqual("{sitecollection}/_catalogs/MasterPage/oslo.master", webSettings.MasterPageUrl);
            Assert.AreEqual("{sitecollection}/_catalogs/MasterPage/CustomMaster.master", webSettings.CustomMasterPageUrl);
            Assert.AreEqual("/sites/hubsite", webSettings.HubSiteUrl);
            Assert.AreEqual(false, webSettings.CommentsOnSitePagesDisabled);
            Assert.AreEqual(true, webSettings.QuickLaunchEnabled);
            Assert.AreEqual(true, webSettings.IsMultilingual);
            Assert.AreEqual(true, webSettings.OverwriteTranslationsOnChange);
            Assert.IsNotNull(webSettings.AlternateUICultures);
            Assert.AreEqual(3, webSettings.AlternateUICultures.Length);
            Assert.AreEqual(1033, webSettings.AlternateUICultures[0].LCID);
            Assert.AreEqual(1040, webSettings.AlternateUICultures[1].LCID);
            Assert.AreEqual(1035, webSettings.AlternateUICultures[2].LCID);
            Assert.AreEqual(true, webSettings.ExcludeFromOfflineClient);
            Assert.AreEqual(true, webSettings.MembersCanShare);
            Assert.AreEqual(false, webSettings.DisableFlows);
            Assert.AreEqual(false, webSettings.DisableAppViews);
            Assert.AreEqual(true, webSettings.HorizontalQuickLaunch);
            Assert.AreEqual(WebSettingsSearchScope.Hub, webSettings.SearchScope);
            Assert.AreEqual(TargetNamespace.SearchBoxInNavBar.AllPages, webSettings.SearchBoxInNavBar);
            Assert.AreEqual("https://contoso.sharepoint.com/sites/SearchCenter", webSettings.SearchCenterUrl);
        }

        [TestMethod]
        [TestCategory(TEST_CATEGORY)]
        public void XMLSerializer_Deserialize_SiteSettings()
        {
            var provider = new XMLFileSystemTemplateProvider($@"{AppDomain.CurrentDomain.BaseDirectory}\..\..\..\Resources", "Templates");

            var serializer = new TargetSerializer();
            var template = provider.GetTemplate(TEST_TEMPLATE, serializer);
            var siteSettings = template.SiteSettings;

            Assert.AreEqual(true, siteSettings.AllowDesigner);
            Assert.AreEqual(false, siteSettings.AllowCreateDeclarativeWorkflow);
            Assert.AreEqual(false, siteSettings.AllowSaveDeclarativeWorkflowAsTemplate);
            Assert.AreEqual(false, siteSettings.AllowSavePublishDeclarativeWorkflow);
            Assert.AreEqual(true, siteSettings.SocialBarOnSitePagesDisabled);
            Assert.AreEqual(PnP.Framework.Provisioning.Model.SearchBoxInNavBar.ModernOnly, siteSettings.SearchBoxInNavBar);
            Assert.AreEqual(true, siteSettings.ShowPeoplePickerSuggestionsForGuestUsers);
            Assert.AreEqual("https://contoso.sharepoint.com/sites/SearchCenter", siteSettings.SearchCenterUrl);
        }

        [TestMethod]
        [TestCategory(TEST_CATEGORY)]
        public void XMLSerializer_Serialize_SiteSettings()
        {
            var provider = new XMLFileSystemTemplateProvider($@"{AppDomain.CurrentDomain.BaseDirectory}\..\..\..\Resources", "Templates");

            var result = new ProvisioningTemplate
            {
                SiteSettings = new PnP.Framework.Provisioning.Model.SiteSettings
                {
                    AllowDesigner = true,
                    AllowCreateDeclarativeWorkflow = false,
                    AllowSaveDeclarativeWorkflowAsTemplate = false,
                    AllowSavePublishDeclarativeWorkflow = false,
                    SocialBarOnSitePagesDisabled = true,
                    SearchBoxInNavBar = PnP.Framework.Provisioning.Model.SearchBoxInNavBar.ModernOnly,
                    ShowPeoplePickerSuggestionsForGuestUsers = true,
                    SearchCenterUrl = "https://contoso.sharepoint.com/sites/SearchCenter",
                }
            };

            var serializer = new TargetSerializer();
            provider.SaveAs(result, TEST_OUT_FILE, serializer);

            var path = $"{provider.Connector.Parameters["ConnectionString"]}\\{provider.Connector.Parameters["Container"]}\\{TEST_OUT_FILE}";
            Assert.IsTrue(File.Exists(path)); path = Path.GetFullPath(path);
            var xml = XDocument.Load(path);
            var wrappedResult = XMLSerializer.Deserialize<TargetProvisioning>(xml);

            var template = wrappedResult.Templates[0].ProvisioningTemplate.First();

            var siteSettings = template.SiteSettings;

            Assert.AreEqual(true, siteSettings.AllowDesigner);
            Assert.AreEqual(false, siteSettings.AllowCreateDeclarativeWorkflow);
            Assert.AreEqual(false, siteSettings.AllowSaveDeclarativeWorkflowAsTemplate);
            Assert.AreEqual(false, siteSettings.AllowSavePublishDeclarativeWorkflow);
            Assert.AreEqual(true, siteSettings.SocialBarOnSitePagesDisabled);
            Assert.AreEqual(TargetNamespace.SearchBoxInNavBar.ModernOnly, siteSettings.SearchBoxInNavBar);
            Assert.AreEqual(true, siteSettings.ShowPeoplePickerSuggestionsForGuestUsers);
            Assert.AreEqual("https://contoso.sharepoint.com/sites/SearchCenter", siteSettings.SearchCenterUrl);
        }

        [TestMethod]
        [TestCategory(TEST_CATEGORY)]
        public void XMLSerializer_Deserialize_RegionalSettings()
        {
            var provider = new XMLFileSystemTemplateProvider($@"{AppDomain.CurrentDomain.BaseDirectory}\..\..\..\Resources", "Templates");

            var serializer = new TargetSerializer();
            var template = provider.GetTemplate(TEST_TEMPLATE, serializer);
            var regionalSettings = template.RegionalSettings;

            Assert.AreEqual(1, regionalSettings.AdjustHijriDays);
            Assert.AreEqual(CalendarType.ChineseLunar, regionalSettings.AlternateCalendarType);
            Assert.AreEqual(CalendarType.Hebrew, regionalSettings.CalendarType);
            Assert.AreEqual(5, regionalSettings.Collation);
            Assert.AreEqual(DayOfWeek.Sunday, regionalSettings.FirstDayOfWeek);
            Assert.AreEqual(1, regionalSettings.FirstWeekOfYear);
            Assert.AreEqual(1040, regionalSettings.LocaleId);
            Assert.AreEqual(true, regionalSettings.ShowWeeks);
            Assert.AreEqual(true, regionalSettings.Time24);
            Assert.AreEqual(4, regionalSettings.TimeZone);
            Assert.AreEqual(WorkHour.PM0500, regionalSettings.WorkDayEndHour);
            Assert.AreEqual(WorkHour.AM0900, regionalSettings.WorkDayStartHour);
            Assert.AreEqual(62, regionalSettings.WorkDays);
        }

        [TestMethod]
        [TestCategory(TEST_CATEGORY)]
        public void XMLSerializer_Serialize_RegionalSettings()
        {
            var provider = new XMLFileSystemTemplateProvider($@"{AppDomain.CurrentDomain.BaseDirectory}\..\..\..\Resources", "Templates");

            var result = new ProvisioningTemplate
            {
                RegionalSettings = new PnP.Framework.Provisioning.Model.RegionalSettings
                {
                    AdjustHijriDays = 1,
                    AlternateCalendarType = CalendarType.ChineseLunar,
                    CalendarType = CalendarType.Hebrew,
                    Collation = 5,
                    FirstDayOfWeek = DayOfWeek.Sunday,
                    FirstWeekOfYear = 1,
                    LocaleId = 1040,
                    ShowWeeks = true,
                    Time24 = true,
                    TimeZone = 4,
                    WorkDayEndHour = WorkHour.PM0500,
                    WorkDayStartHour = WorkHour.AM0900,
                    WorkDays = 62
                }
            };

            var serializer = new TargetSerializer();
            provider.SaveAs(result, TEST_OUT_FILE, serializer);

            var path = $"{provider.Connector.Parameters["ConnectionString"]}\\{provider.Connector.Parameters["Container"]}\\{TEST_OUT_FILE}";
            Assert.IsTrue(File.Exists(path)); path = Path.GetFullPath(path);
            var xml = XDocument.Load(path);
            var wrappedResult = XMLSerializer.Deserialize<TargetProvisioning>(xml);

            var template = wrappedResult.Templates[0].ProvisioningTemplate.First();

            var regionalSettings = template.RegionalSettings;

            Assert.AreEqual(1, regionalSettings.AdjustHijriDays);
            Assert.AreEqual(TargetNamespace.CalendarType.ChineseLunar, regionalSettings.AlternateCalendarType);
            Assert.AreEqual(TargetNamespace.CalendarType.Hebrew, regionalSettings.CalendarType);
            Assert.AreEqual(5, regionalSettings.Collation);
            Assert.AreEqual(TargetNamespace.DayOfWeek.Sunday, regionalSettings.FirstDayOfWeek);
            Assert.AreEqual(1, regionalSettings.FirstWeekOfYear);
            Assert.AreEqual(1040, regionalSettings.LocaleId);
            Assert.AreEqual(true, regionalSettings.ShowWeeks);
            Assert.AreEqual(true, regionalSettings.Time24);
            Assert.AreEqual("4", regionalSettings.TimeZone);
            Assert.AreEqual(TargetNamespace.WorkHour.Item500PM, regionalSettings.WorkDayEndHour);
            Assert.AreEqual(TargetNamespace.WorkHour.Item900AM, regionalSettings.WorkDayStartHour);
            Assert.AreEqual(62, regionalSettings.WorkDays);
        }

        [TestMethod]
        [TestCategory(TEST_CATEGORY)]
        public void XMLSerializer_Deserialize_SupportedUILanguages()
        {
            var provider = new XMLFileSystemTemplateProvider($@"{AppDomain.CurrentDomain.BaseDirectory}\..\..\..\Resources", "Templates");

            var serializer = new TargetSerializer();
            var template = provider.GetTemplate(TEST_TEMPLATE, serializer);
            var supportedUiLanguages = template.SupportedUILanguages;

            Assert.AreEqual(3, supportedUiLanguages.Count);
            Assert.AreEqual(1033, supportedUiLanguages[0].LCID);
        }

        [TestMethod]
        [TestCategory(TEST_CATEGORY)]
        public void XMLSerializer_Serialize_SupportedUILanguages()
        {
            var provider = new XMLFileSystemTemplateProvider($@"{AppDomain.CurrentDomain.BaseDirectory}\..\..\..\Resources", "Templates");

            var result = new ProvisioningTemplate();
            result.SupportedUILanguages.Add(new SupportedUILanguage
            {
                LCID = 1033
            });

            var serializer = new TargetSerializer();
            provider.SaveAs(result, TEST_OUT_FILE, serializer);

            var path = $"{provider.Connector.Parameters["ConnectionString"]}\\{provider.Connector.Parameters["Container"]}\\{TEST_OUT_FILE}";
            Assert.IsTrue(File.Exists(path)); path = Path.GetFullPath(path);
            var xml = XDocument.Load(path);
            var wrappedResult = XMLSerializer.Deserialize<TargetProvisioning>(xml);

            var template = wrappedResult.Templates[0].ProvisioningTemplate.First();

            var supportedUiLanguages = template.SupportedUILanguages;

            Assert.AreEqual(1, supportedUiLanguages.Count());
            Assert.AreEqual(1033, supportedUiLanguages[0].LCID);
        }

        [TestMethod]
        [TestCategory(TEST_CATEGORY)]
        public void XMLSerializer_Deserialize_AuditSettings()
        {
            var provider = new XMLFileSystemTemplateProvider($@"{AppDomain.CurrentDomain.BaseDirectory}\..\..\..\Resources", "Templates");

            var serializer = new TargetSerializer();
            var template = provider.GetTemplate(TEST_TEMPLATE, serializer);
            var auditSettings = template.AuditSettings;

            Assert.AreEqual(50, auditSettings.AuditLogTrimmingRetention);
            Assert.AreEqual(true, auditSettings.TrimAuditLog);
            Assert.AreEqual(AuditMaskType.CheckIn | AuditMaskType.CheckOut | AuditMaskType.Update | AuditMaskType.View, auditSettings.AuditFlags);
        }

        [TestMethod]
        [TestCategory(TEST_CATEGORY)]
        public void XMLSerializer_Serialize_AuditSettings()
        {
            var provider = new XMLFileSystemTemplateProvider($@"{AppDomain.CurrentDomain.BaseDirectory}\..\..\..\Resources", "Templates");

            var result = new ProvisioningTemplate
            {
                AuditSettings = new PnP.Framework.Provisioning.Model.AuditSettings
                {
                    TrimAuditLog = true,
                    AuditLogTrimmingRetention = 50,
                    AuditFlags = AuditMaskType.CheckIn | AuditMaskType.CheckOut | AuditMaskType.Update |
                                 AuditMaskType.View
                }
            };

            var serializer = new TargetSerializer();
            provider.SaveAs(result, TEST_OUT_FILE, serializer);

            var path = $"{provider.Connector.Parameters["ConnectionString"]}\\{provider.Connector.Parameters["Container"]}\\{TEST_OUT_FILE}";
            Assert.IsTrue(File.Exists(path)); path = Path.GetFullPath(path);
            var xml = XDocument.Load(path);
            var wrappedResult = XMLSerializer.Deserialize<TargetProvisioning>(xml);

            var template = wrappedResult.Templates[0].ProvisioningTemplate.First();

            var auditSettings = template.AuditSettings;

            Assert.AreEqual(50, auditSettings.AuditLogTrimmingRetention);
            Assert.AreEqual(true, auditSettings.TrimAuditLog);
            Assert.IsTrue(auditSettings.Audit.SingleOrDefault(a => a.AuditFlag == AuditSettingsAuditAuditFlag.CheckIn) != null);
            Assert.IsTrue(auditSettings.Audit.SingleOrDefault(a => a.AuditFlag == AuditSettingsAuditAuditFlag.CheckOut) != null);
            Assert.IsTrue(auditSettings.Audit.SingleOrDefault(a => a.AuditFlag == AuditSettingsAuditAuditFlag.Update) != null);
            Assert.IsTrue(auditSettings.Audit.SingleOrDefault(a => a.AuditFlag == AuditSettingsAuditAuditFlag.View) != null);
        }

        [TestMethod]
        [TestCategory(TEST_CATEGORY)]
        public void XMLSerializer_Deserialize_PropertyBagEntries()
        {
            var provider = new XMLFileSystemTemplateProvider($@"{AppDomain.CurrentDomain.BaseDirectory}\..\..\..\Resources", "Templates");

            var serializer = new TargetSerializer();
            var template = provider.GetTemplate(TEST_TEMPLATE, serializer);
            var propertyBagEntries = template.PropertyBagEntries;

            Assert.AreEqual(true, propertyBagEntries[0].Overwrite);
            Assert.AreEqual("KEY1", propertyBagEntries[0].Key);
            Assert.AreEqual("value1", propertyBagEntries[0].Value);
            Assert.AreEqual(true, propertyBagEntries[1].Indexed);
        }

        [TestMethod]
        [TestCategory(TEST_CATEGORY)]
        public void XMLSerializer_Serialize_PropertyBagEntries()
        {
            var provider = new XMLFileSystemTemplateProvider($@"{AppDomain.CurrentDomain.BaseDirectory}\..\..\..\Resources", "Templates");

            var result = new ProvisioningTemplate();
            result.PropertyBagEntries.Add(new PnP.Framework.Provisioning.Model.PropertyBagEntry
            {
                Overwrite = true,
                Key = "KEY1",
                Value = "value1"
            });
            result.PropertyBagEntries.Add(new PnP.Framework.Provisioning.Model.PropertyBagEntry
            {
                Indexed = true,
                Key = "KEY2",
                Value = "value2"
            });

            var serializer = new TargetSerializer();
            provider.SaveAs(result, TEST_OUT_FILE, serializer);

            var path = $"{provider.Connector.Parameters["ConnectionString"]}\\{provider.Connector.Parameters["Container"]}\\{TEST_OUT_FILE}";
            Assert.IsTrue(File.Exists(path)); path = Path.GetFullPath(path);
            var xml = XDocument.Load(path);
            var wrappedResult = XMLSerializer.Deserialize<TargetProvisioning>(xml);

            var template = wrappedResult.Templates[0].ProvisioningTemplate.First();

            var propertyBagEntries = template.PropertyBagEntries;

            Assert.AreEqual(true, propertyBagEntries[0].Overwrite);
            Assert.AreEqual("KEY1", propertyBagEntries[0].Key);
            Assert.AreEqual("value1", propertyBagEntries[0].Value);
            Assert.AreEqual(true, propertyBagEntries[1].Indexed);
        }

        [TestMethod]
        [TestCategory(TEST_CATEGORY)]
        public void XMLSerializer_Deserialize_Security()
        {
            var provider = new XMLFileSystemTemplateProvider($@"{AppDomain.CurrentDomain.BaseDirectory}\..\..\..\Resources", "Templates");

            var serializer = new TargetSerializer();
            var template = provider.GetTemplate(TEST_TEMPLATE, serializer);
            var security = template.Security;

            // security common properties
            Assert.IsNotNull(security);
            Assert.IsTrue(security.BreakRoleInheritance);
            Assert.IsTrue(security.ClearSubscopes);
            Assert.IsFalse(security.CopyRoleAssignments);
            Assert.AreEqual("Test Value", security.AssociatedGroups);
            Assert.AreEqual("{parameter:AssociatedMemberGroup}", security.AssociatedMemberGroup);
            Assert.AreEqual("{parameter:AssociatedOwnerGroup}", security.AssociatedOwnerGroup);
            Assert.AreEqual("{parameter:AssociatedVisitorGroup}", security.AssociatedVisitorGroup);
            Assert.AreEqual(true, security.RemoveExistingUniqueRoleAssignments);
            Assert.AreEqual(true, security.ResetRoleInheritance);

            // additional administrators
            Assert.IsNotNull(security.AdditionalAdministrators);
            Assert.AreEqual(2, security.AdditionalAdministrators.Count);
            Assert.IsNotNull(security.AdditionalAdministrators.FirstOrDefault(u => u.Name == "user@contoso.com"));
            Assert.IsNotNull(security.AdditionalAdministrators.FirstOrDefault(u => u.Name == "U_SHAREPOINT_ADMINS"));
            Assert.IsFalse(security.ClearExistingAdministrators);

            // additional owners
            Assert.IsNotNull(security.AdditionalOwners);
            Assert.AreEqual(2, security.AdditionalOwners.Count);
            Assert.IsNotNull(security.AdditionalOwners.FirstOrDefault(u => u.Name == "user@contoso.com"));
            Assert.IsNotNull(security.AdditionalOwners.FirstOrDefault(u => u.Name == "U_SHAREPOINT_ADMINS"));
            Assert.IsTrue(security.ClearExistingOwners);

            // additional members
            Assert.IsNotNull(security.AdditionalMembers);
            Assert.AreEqual(2, security.AdditionalMembers.Count);
            Assert.IsNotNull(security.AdditionalMembers.FirstOrDefault(u => u.Name == "user@contoso.com"));
            Assert.IsNotNull(security.AdditionalMembers.FirstOrDefault(u => u.Name == "U_SHAREPOINT_ADMINS"));
            Assert.IsFalse(security.ClearExistingMembers);

            // additional visitors
            Assert.IsNotNull(security.AdditionalVisitors);
            Assert.AreEqual(2, security.AdditionalVisitors.Count);
            Assert.IsNotNull(security.AdditionalVisitors.FirstOrDefault(u => u.Name == "user@contoso.com"));
            Assert.IsNotNull(security.AdditionalVisitors.FirstOrDefault(u => u.Name == "U_SHAREPOINT_ADMINS"));

            // permissions
            Assert.IsNotNull(security.SiteSecurityPermissions);
            Assert.IsNotNull(security.SiteSecurityPermissions.RoleDefinitions);
            Assert.AreEqual(1, security.SiteSecurityPermissions.RoleDefinitions.Count);
            var role = security.SiteSecurityPermissions.RoleDefinitions.FirstOrDefault(r => r.Name == "Manage List Items");
            Assert.IsNotNull(role);
            Assert.AreEqual("Allows a user to manage list items", role.Description);
            Assert.IsNotNull(role.Permissions);
            Assert.AreEqual(4, role.Permissions.Count);
            Assert.IsTrue(role.Permissions.Contains(PermissionKind.ViewListItems));
            Assert.IsTrue(role.Permissions.Contains(PermissionKind.AddListItems));
            Assert.IsTrue(role.Permissions.Contains(PermissionKind.EditListItems));
            Assert.IsTrue(role.Permissions.Contains(PermissionKind.DeleteListItems));

            Assert.IsNotNull(security.SiteSecurityPermissions.RoleAssignments);
            Assert.AreEqual(4, security.SiteSecurityPermissions.RoleAssignments.Count);

            // role assignments
            var assign = security.SiteSecurityPermissions.RoleAssignments.FirstOrDefault(p => p.Principal == "user1@contoso.com");
            Assert.IsNotNull(assign);
            Assert.AreEqual("Manage List Items", assign.RoleDefinition);

            Assert.IsNotNull(security.SiteGroups);
            Assert.AreEqual(2, security.SiteGroups.Count);

            // site groups
            var group = security.SiteGroups.FirstOrDefault(g => g.Title == "TestGroup1");
            Assert.IsNotNull(group);
            Assert.AreEqual("Test Group 1", group.Description);
            Assert.AreEqual("user1@contoso.com", group.Owner);
            Assert.AreEqual("group1@contoso.com", group.RequestToJoinLeaveEmailSetting);
            Assert.IsTrue(group.AllowMembersEditMembership);
            Assert.IsTrue(group.AllowRequestToJoinLeave);
            Assert.IsTrue(group.AutoAcceptRequestToJoinLeave);
            Assert.IsTrue(group.OnlyAllowMembersViewMembership);
            Assert.IsNotNull(group.Members);
            Assert.AreEqual(2, group.Members.Count);
            Assert.IsNotNull(group.Members.FirstOrDefault(m => m.Name == "user1@contoso.com"));
            Assert.IsNotNull(group.Members.FirstOrDefault(m => m.Name == "user2@contoso.com"));
            Assert.IsFalse(group.ClearExistingMembers);

            group = security.SiteGroups.FirstOrDefault(g => g.Title == "Power Users");
            Assert.IsNotNull(group);
            Assert.AreEqual("admin@contoso.com", group.Owner);
            Assert.IsTrue(string.IsNullOrEmpty(group.RequestToJoinLeaveEmailSetting));
            Assert.IsFalse(group.AllowMembersEditMembership);
            Assert.IsFalse(group.AllowRequestToJoinLeave);
            Assert.IsFalse(group.AutoAcceptRequestToJoinLeave);
            Assert.IsFalse(group.OnlyAllowMembersViewMembership);
            Assert.AreEqual(3, group.Members.Count);
            Assert.IsTrue(group.ClearExistingMembers);
        }

        [TestMethod]
        [TestCategory(TEST_CATEGORY)]
        public void XMLSerializer_Serialize_Security()
        {
            var provider = new XMLFileSystemTemplateProvider($@"{AppDomain.CurrentDomain.BaseDirectory}\..\..\..\Resources", "Templates");

            var result = new ProvisioningTemplate
            {
                Security = new SiteSecurity()
                {
                    BreakRoleInheritance = true,
                    ClearSubscopes = true,
                    CopyRoleAssignments = true,
                    AssociatedGroups = "Test Value",
                    AssociatedMemberGroup = "{parameter:AssociatedMemberGroup}",
                    AssociatedOwnerGroup = "{parameter:AssociatedOwnerGroup}",
                    AssociatedVisitorGroup = "{parameter:AssociatedVisitorGroup}",
                    ClearExistingAdministrators = true,
                    ClearExistingVisitors = true,
                    ResetRoleInheritance = true,
                    RemoveExistingUniqueRoleAssignments = true
                }
            };
            result.Security.AdditionalAdministrators.Add(new PnP.Framework.Provisioning.Model.User() { Name = "user@contoso.com" });
            result.Security.AdditionalAdministrators.Add(new PnP.Framework.Provisioning.Model.User() { Name = "U_SHAREPOINT_ADMINS" });
            result.Security.AdditionalOwners.Add(new PnP.Framework.Provisioning.Model.User() { Name = "user@contoso.com" });
            result.Security.AdditionalOwners.Add(new PnP.Framework.Provisioning.Model.User() { Name = "U_SHAREPOINT_ADMINS" });
            result.Security.AdditionalMembers.Add(new PnP.Framework.Provisioning.Model.User() { Name = "user@contoso.com" });
            result.Security.AdditionalMembers.Add(new PnP.Framework.Provisioning.Model.User() { Name = "U_SHAREPOINT_ADMINS" });
            result.Security.AdditionalVisitors.Add(new PnP.Framework.Provisioning.Model.User() { Name = "user@contoso.com" });
            result.Security.AdditionalVisitors.Add(new PnP.Framework.Provisioning.Model.User() { Name = "U_SHAREPOINT_ADMINS" });

            result.Security.SiteSecurityPermissions.RoleDefinitions.Add(new PnP.Framework.Provisioning.Model.RoleDefinition(new List<PermissionKind> {
                PermissionKind.ViewListItems,
                PermissionKind.AddListItems
            })
            {
                Name = "User",
                Description = "User Role"
            });
            result.Security.SiteSecurityPermissions.RoleDefinitions.Add(new PnP.Framework.Provisioning.Model.RoleDefinition(new List<PermissionKind> {
                PermissionKind.EmptyMask
            })
            {
                Name = "EmptyRole",
                Description = "Empty Role"
            });
            result.Security.SiteSecurityPermissions.RoleAssignments.Add(new PnP.Framework.Provisioning.Model.RoleAssignment
            {
                Principal = "admin@contoso.com",
                RoleDefinition = "Owner"
            });
            result.Security.SiteSecurityPermissions.RoleAssignments.Add(new PnP.Framework.Provisioning.Model.RoleAssignment
            {
                Principal = "user@contoso.com",
                RoleDefinition = "User"
            });

            result.Security.SiteGroups.Add(new PnP.Framework.Provisioning.Model.SiteGroup(new List<PnP.Framework.Provisioning.Model.User>
            {
                new PnP.Framework.Provisioning.Model.User
                {
                     Name = "user1@contoso.com"
                },
                new PnP.Framework.Provisioning.Model.User
                {
                     Name = "user2@contoso.com"
                }
            })
            {
                AllowMembersEditMembership = true,
                AllowRequestToJoinLeave = true,
                AutoAcceptRequestToJoinLeave = true,
                Description = "Test Group 1",
                OnlyAllowMembersViewMembership = true,
                Owner = "user1@contoso.com",
                RequestToJoinLeaveEmailSetting = "group1@contoso.com",
                Title = "TestGroup1",
                ClearExistingMembers = true
            });
            result.Security.SiteGroups.Add(new PnP.Framework.Provisioning.Model.SiteGroup(new List<PnP.Framework.Provisioning.Model.User>
            {
                new PnP.Framework.Provisioning.Model.User
                {
                    Name = "user1@contoso.com"
                }
            })
            {
                Title = "TestGroup2",
                Owner = "user2@contoso.com"
            });

            var serializer = new TargetSerializer();
            provider.SaveAs(result, TEST_OUT_FILE, serializer);

            var path = $"{provider.Connector.Parameters["ConnectionString"]}\\{provider.Connector.Parameters["Container"]}\\{TEST_OUT_FILE}";
            Assert.IsTrue(File.Exists(path)); path = Path.GetFullPath(path);
            var xml = XDocument.Load(path);
            var wrappedResult = XMLSerializer.Deserialize<TargetProvisioning>(xml);

            var template = wrappedResult.Templates[0].ProvisioningTemplate.First();

            //security common properties
            Assert.IsNotNull(template.Security);
            Assert.IsTrue(template.Security.BreakRoleInheritance);
            Assert.IsTrue(template.Security.ClearSubscopes);
            Assert.IsTrue(template.Security.CopyRoleAssignments);
            Assert.AreEqual("Test Value", template.Security.AssociatedGroups);
            Assert.AreEqual("{parameter:AssociatedMemberGroup}", template.Security.AssociatedMemberGroup);
            Assert.AreEqual("{parameter:AssociatedOwnerGroup}", template.Security.AssociatedOwnerGroup);
            Assert.AreEqual("{parameter:AssociatedVisitorGroup}", template.Security.AssociatedVisitorGroup);
            Assert.AreEqual(true, template.Security.RemoveExistingUniqueRoleAssignments);
            Assert.AreEqual(true, template.Security.ResetRoleInheritance);

            // additional adminstrators
            Assert.IsNotNull(template.Security.AdditionalAdministrators);
            Assert.AreEqual(2, template.Security.AdditionalAdministrators.User.Length);
            Assert.IsNotNull(template.Security.AdditionalAdministrators.User.FirstOrDefault(u => u.Name == "user@contoso.com"));
            Assert.IsNotNull(template.Security.AdditionalAdministrators.User.FirstOrDefault(u => u.Name == "U_SHAREPOINT_ADMINS"));
            Assert.IsTrue(template.Security.AdditionalAdministrators.ClearExistingItems);

            // additional owners
            Assert.IsNotNull(template.Security.AdditionalOwners);
            Assert.AreEqual(2, template.Security.AdditionalOwners.User.Length);
            Assert.IsNotNull(template.Security.AdditionalOwners.User.FirstOrDefault(u => u.Name == "user@contoso.com"));
            Assert.IsNotNull(template.Security.AdditionalOwners.User.FirstOrDefault(u => u.Name == "U_SHAREPOINT_ADMINS"));

            // additional members
            Assert.IsNotNull(template.Security.AdditionalMembers);
            Assert.AreEqual(2, template.Security.AdditionalMembers.User.Length);
            Assert.IsNotNull(template.Security.AdditionalMembers.User.FirstOrDefault(u => u.Name == "user@contoso.com"));
            Assert.IsNotNull(template.Security.AdditionalMembers.User.FirstOrDefault(u => u.Name == "U_SHAREPOINT_ADMINS"));
            Assert.IsFalse(template.Security.AdditionalMembers.ClearExistingItems);

            // additional visitors
            Assert.IsNotNull(template.Security.AdditionalVisitors);
            Assert.AreEqual(2, template.Security.AdditionalVisitors.User.Length);
            Assert.IsNotNull(template.Security.AdditionalVisitors.User.FirstOrDefault(u => u.Name == "user@contoso.com"));
            Assert.IsNotNull(template.Security.AdditionalVisitors.User.FirstOrDefault(u => u.Name == "U_SHAREPOINT_ADMINS"));
            Assert.IsTrue(template.Security.AdditionalVisitors.ClearExistingItems);

            // permissions
            Assert.IsNotNull(template.Security.Permissions);
            Assert.IsNotNull(template.Security.Permissions.RoleDefinitions);
            Assert.AreEqual(2, template.Security.Permissions.RoleDefinitions.Length);
            var role = template.Security.Permissions.RoleDefinitions.FirstOrDefault(r => r.Name == "User");
            Assert.IsNotNull(role);
            Assert.AreEqual("User Role", role.Description);
            Assert.IsNotNull(role.Permissions);
            Assert.AreEqual(2, role.Permissions.Length);
            Assert.IsTrue(role.Permissions.Contains(RoleDefinitionPermission.ViewListItems));
            Assert.IsTrue(role.Permissions.Contains(RoleDefinitionPermission.AddListItems));

            role = template.Security.Permissions.RoleDefinitions.FirstOrDefault(r => r.Name == "EmptyRole");
            Assert.IsNotNull(role);
            Assert.AreEqual("Empty Role", role.Description);
            Assert.IsNotNull(role.Permissions);
            Assert.AreEqual(1, role.Permissions.Length);
            Assert.IsTrue(role.Permissions.Contains(RoleDefinitionPermission.EmptyMask));

            Assert.IsNotNull(template.Security.Permissions);
            Assert.IsNotNull(template.Security.Permissions.RoleAssignments);
            Assert.AreEqual(2, template.Security.Permissions.RoleAssignments.Length);
            var assign = template.Security.Permissions.RoleAssignments.FirstOrDefault(p => p.Principal == "admin@contoso.com");
            Assert.IsNotNull(assign);
            Assert.AreEqual("Owner", assign.RoleDefinition);
            assign = template.Security.Permissions.RoleAssignments.FirstOrDefault(p => p.Principal == "user@contoso.com");
            Assert.IsNotNull(assign);
            Assert.AreEqual("User", assign.RoleDefinition);

            // site groups
            Assert.IsNotNull(template.Security.SiteGroups);
            Assert.AreEqual(2, template.Security.SiteGroups.Length);
            var group = template.Security.SiteGroups.FirstOrDefault(g => g.Title == "TestGroup1");
            Assert.IsNotNull(group);
            Assert.AreEqual("Test Group 1", group.Description);
            Assert.AreEqual("user1@contoso.com", group.Owner);
            Assert.AreEqual("group1@contoso.com", group.RequestToJoinLeaveEmailSetting);
            Assert.IsTrue(group.AllowMembersEditMembership);
            Assert.IsTrue(group.AllowMembersEditMembershipSpecified);
            Assert.IsTrue(group.AllowRequestToJoinLeave);
            Assert.IsTrue(group.AllowRequestToJoinLeaveSpecified);
            Assert.IsTrue(group.AutoAcceptRequestToJoinLeave);
            Assert.IsTrue(group.AutoAcceptRequestToJoinLeaveSpecified);
            Assert.IsTrue(group.OnlyAllowMembersViewMembership);
            Assert.IsTrue(group.OnlyAllowMembersViewMembershipSpecified);
            Assert.IsNotNull(group.Members);
            Assert.AreEqual(2, group.Members.User.Length);
            Assert.IsNotNull(group.Members.User.FirstOrDefault(m => m.Name == "user1@contoso.com"));
            Assert.IsNotNull(group.Members.User.FirstOrDefault(m => m.Name == "user2@contoso.com"));
            Assert.IsTrue(group.Members.ClearExistingItems);

            group = template.Security.SiteGroups.FirstOrDefault(g => g.Title == "TestGroup2");
            Assert.IsNotNull(group);
            Assert.AreEqual("user2@contoso.com", group.Owner);
            Assert.IsTrue(string.IsNullOrEmpty(group.Description));
            Assert.IsTrue(string.IsNullOrEmpty(group.RequestToJoinLeaveEmailSetting));
            Assert.IsFalse(group.AllowMembersEditMembership);
            Assert.IsFalse(group.AllowRequestToJoinLeave);
            Assert.IsFalse(group.AutoAcceptRequestToJoinLeave);
            Assert.IsFalse(group.OnlyAllowMembersViewMembership);
            Assert.IsFalse(group.Members.ClearExistingItems);
        }

        [TestMethod]
        [TestCategory(TEST_CATEGORY)]
        public void XMLSerializer_Deserialize_Navigation()
        {
            var provider = new XMLFileSystemTemplateProvider($@"{AppDomain.CurrentDomain.BaseDirectory}\..\..\..\Resources", "Templates");

            var serializer = new TargetSerializer();
            var template = provider.GetTemplate(TEST_TEMPLATE, serializer);

            // common properties
            Assert.IsNotNull(template.Navigation);
            Assert.AreEqual(true, template.Navigation.EnableTreeView);
            Assert.AreEqual(true, template.Navigation.AddNewPagesToNavigation);
            Assert.AreEqual(true, template.Navigation.CreateFriendlyUrlsForNewPages);
            Assert.AreEqual(false, template.Navigation.EnableAudienceTargeting);

            // global navigation
            Assert.IsNotNull(template.Navigation.GlobalNavigation);
            Assert.AreEqual(GlobalNavigationType.Managed, template.Navigation.GlobalNavigation.NavigationType);
            Assert.IsNull(template.Navigation.GlobalNavigation.StructuralNavigation);
            Assert.IsNotNull(template.Navigation.GlobalNavigation.ManagedNavigation);
            Assert.AreEqual("{sitecollectionnavigationtermsetid}", template.Navigation.GlobalNavigation.ManagedNavigation.TermSetId);
            Assert.AreEqual("{sitecollectiontermstoreid}", template.Navigation.GlobalNavigation.ManagedNavigation.TermStoreId);

            // current navigation
            Assert.IsNotNull(template.Navigation.CurrentNavigation);
            Assert.AreEqual(CurrentNavigationType.Structural, template.Navigation.CurrentNavigation.NavigationType);
            Assert.IsNull(template.Navigation.CurrentNavigation.ManagedNavigation);
            Assert.IsNotNull(template.Navigation.CurrentNavigation.StructuralNavigation);
            Assert.IsTrue(template.Navigation.CurrentNavigation.StructuralNavigation.RemoveExistingNodes);
            Assert.IsNotNull(template.Navigation.CurrentNavigation.StructuralNavigation.NavigationNodes);
            Assert.AreEqual(3, template.Navigation.CurrentNavigation.StructuralNavigation.NavigationNodes.Count);

            var homeNode = template.Navigation.CurrentNavigation.StructuralNavigation.NavigationNodes.FirstOrDefault(n => n.Title == "Home");
            Assert.IsNotNull(homeNode);
            Assert.AreEqual("Default.aspx", homeNode.Url);
            Assert.IsFalse(homeNode.IsExternal);
            Assert.AreEqual(0, homeNode.NavigationNodes.Count);

            var librariesNode = template.Navigation.CurrentNavigation.StructuralNavigation.NavigationNodes.FirstOrDefault(n => n.Title == "Libraries");
            Assert.IsNotNull(librariesNode);
            Assert.IsTrue(string.IsNullOrEmpty(librariesNode.Url));
            Assert.IsFalse(librariesNode.IsExternal);
            Assert.IsNotNull(librariesNode.NavigationNodes);
            Assert.AreEqual(2, librariesNode.NavigationNodes.Count);

            var invoicesNode = librariesNode.NavigationNodes.FirstOrDefault(n => n.Title == "Invoices");
            Assert.IsNotNull(invoicesNode);
            Assert.AreEqual("Invoices/Forms/AllItems.aspx", invoicesNode.Url);
            Assert.IsFalse(invoicesNode.IsExternal);
            Assert.AreEqual(0, invoicesNode.NavigationNodes.Count);

            var privacyNode = template.Navigation.CurrentNavigation.StructuralNavigation.NavigationNodes.FirstOrDefault(n => n.Title == "Privacy");
            Assert.IsNotNull(privacyNode);
            Assert.AreEqual("http://www.company.com/privacy/", privacyNode.Url);
            Assert.IsTrue(privacyNode.IsExternal);
            Assert.IsNotNull(privacyNode.NavigationNodes);
            Assert.AreEqual(0, privacyNode.NavigationNodes.Count);

            // search navigation
            Assert.AreEqual(false, template.Navigation.SearchNavigation.RemoveExistingNodes);
            Assert.AreEqual("Sample Search Node 01", template.Navigation.SearchNavigation.NavigationNodes[0].Title);
            Assert.AreEqual(true, template.Navigation.SearchNavigation.NavigationNodes[0].IsExternal);
            Assert.AreEqual(false, template.Navigation.SearchNavigation.NavigationNodes[1].IsExternal);
            Assert.AreEqual("{sitecollection}/pages/search01.aspx", template.Navigation.SearchNavigation.NavigationNodes[0].Url);
        }

        [TestMethod]
        [TestCategory(TEST_CATEGORY)]
        public void XMLSerializer_Serialize_Navigation()
        {
            var provider = new XMLFileSystemTemplateProvider($@"{AppDomain.CurrentDomain.BaseDirectory}\..\..\..\Resources", "Templates");

            var result = new ProvisioningTemplate();
            var searchNavigation = new PnP.Framework.Provisioning.Model.StructuralNavigation
            {
                RemoveExistingNodes = false
            };

            searchNavigation.NavigationNodes.Add(new PnP.Framework.Provisioning.Model.NavigationNode
            {
                Title = "Sample Search Node 01",
                IsExternal = true,
                Url = "{sitecollection}/pages/search01.aspx"
            });

            searchNavigation.NavigationNodes.Add(new PnP.Framework.Provisioning.Model.NavigationNode
            {
                Title = "Sample Search Node 02",
                IsExternal = false,
                Url = "{sitecollection}/pages/search02.aspx"
            });

            result.Navigation = new PnP.Framework.Provisioning.Model.Navigation(
                 new GlobalNavigation(GlobalNavigationType.Managed, null, new PnP.Framework.Provisioning.Model.ManagedNavigation()),
                 new CurrentNavigation(CurrentNavigationType.Structural, new PnP.Framework.Provisioning.Model.StructuralNavigation(), null), searchNavigation)
            {
                EnableTreeView = true,
                AddNewPagesToNavigation = true,
                CreateFriendlyUrlsForNewPages = true,
                EnableAudienceTargeting = false
            };

            result.Navigation.GlobalNavigation.ManagedNavigation.TermSetId = "415185a1-ee1c-4ce9-9e38-cea3f854e802";
            result.Navigation.GlobalNavigation.ManagedNavigation.TermStoreId = "c1175ad1-c710-4131-a6c9-aa854a5cc4c4";

            result.Navigation.CurrentNavigation.StructuralNavigation.RemoveExistingNodes = true;
            var node1 = new PnP.Framework.Provisioning.Model.NavigationNode()
            {
                IsExternal = false,
                Title = "Node 1",
                Url = "/Node1.aspx",

            };
            var node11 = new PnP.Framework.Provisioning.Model.NavigationNode()
            {
                IsExternal = true,
                Title = "Node 1.1",
                Url = "http://aka.ms/SharePointPnP"
            };
            var node111 = new PnP.Framework.Provisioning.Model.NavigationNode()
            {
                IsExternal = true,
                Title = "Node 1.1.1",
                Url = "http://aka.ms/OfficeDevPnP"
            };
            var node12 = new PnP.Framework.Provisioning.Model.NavigationNode()
            {
                IsExternal = true,
                Title = "Node 1.2",
                Url = "/Node1-2.aspx"
            };
            var node2 = new PnP.Framework.Provisioning.Model.NavigationNode()
            {
                IsExternal = false,
                Title = "Node 2",
                Url = "/Node1.aspx"
            };
            node11.NavigationNodes.Add(node111);
            node1.NavigationNodes.Add(node11);
            node1.NavigationNodes.Add(node12);
            result.Navigation.CurrentNavigation.StructuralNavigation.NavigationNodes.Add(node1);
            result.Navigation.CurrentNavigation.StructuralNavigation.NavigationNodes.Add(node2);

            var serializer = new TargetSerializer();
            provider.SaveAs(result, TEST_OUT_FILE, serializer);

            var path = $"{provider.Connector.Parameters["ConnectionString"]}\\{provider.Connector.Parameters["Container"]}\\{TEST_OUT_FILE}";
            Assert.IsTrue(File.Exists(path)); path = Path.GetFullPath(path);
            var xml = XDocument.Load(path);
            var wrappedResult = XMLSerializer.Deserialize<TargetProvisioning>(xml);

            var template = wrappedResult.Templates[0].ProvisioningTemplate.First();

            Assert.IsNotNull(template.Navigation);
            Assert.AreEqual(true, template.Navigation.EnableTreeView);
            Assert.AreEqual(true, template.Navigation.AddNewPagesToNavigation);
            Assert.AreEqual(true, template.Navigation.CreateFriendlyUrlsForNewPages);
            Assert.AreEqual(false, template.Navigation.EnableAudienceTargeting);

            Assert.IsNotNull(template.Navigation.GlobalNavigation);
            Assert.AreEqual(NavigationGlobalNavigationNavigationType.Managed, template.Navigation.GlobalNavigation.NavigationType);
            Assert.IsNull(template.Navigation.GlobalNavigation.StructuralNavigation);
            Assert.IsNotNull(template.Navigation.GlobalNavigation.ManagedNavigation);
            Assert.AreEqual("415185a1-ee1c-4ce9-9e38-cea3f854e802", template.Navigation.GlobalNavigation.ManagedNavigation.TermSetId);
            Assert.AreEqual("c1175ad1-c710-4131-a6c9-aa854a5cc4c4", template.Navigation.GlobalNavigation.ManagedNavigation.TermStoreId);

            Assert.IsNotNull(template.Navigation.CurrentNavigation);
            Assert.AreEqual(NavigationCurrentNavigationNavigationType.Structural, template.Navigation.CurrentNavigation.NavigationType);
            Assert.IsNull(template.Navigation.CurrentNavigation.ManagedNavigation);
            Assert.IsNotNull(template.Navigation.CurrentNavigation.StructuralNavigation);
            Assert.IsTrue(template.Navigation.CurrentNavigation.StructuralNavigation.RemoveExistingNodes);
            Assert.IsNotNull(template.Navigation.CurrentNavigation.StructuralNavigation.NavigationNode);
            Assert.AreEqual(2, template.Navigation.CurrentNavigation.StructuralNavigation.NavigationNode.Length);

            var n1 = template.Navigation.CurrentNavigation.StructuralNavigation.NavigationNode.FirstOrDefault(n => n.Title == "Node 1");
            Assert.IsNotNull(n1);
            Assert.AreEqual("/Node1.aspx", n1.Url);
            Assert.IsFalse(n1.IsExternal);
            Assert.IsNotNull(n1.NavigationNode1);
            Assert.AreEqual(2, n1.NavigationNode1.Length);

            var n11 = n1.NavigationNode1.FirstOrDefault(n => n.Title == "Node 1.1");
            Assert.IsNotNull(n11);
            Assert.AreEqual("http://aka.ms/SharePointPnP", n11.Url);
            Assert.IsTrue(n11.IsExternal);
            Assert.IsNotNull(n11.NavigationNode1);
            Assert.AreEqual(1, n11.NavigationNode1.Length);

            var n111 = n11.NavigationNode1.FirstOrDefault(n => n.Title == "Node 1.1.1");
            Assert.IsNotNull(n111);
            Assert.AreEqual("http://aka.ms/OfficeDevPnP", n111.Url);
            Assert.IsTrue(n111.IsExternal);
            Assert.IsNull(n111.NavigationNode1);

            var n12 = n1.NavigationNode1.FirstOrDefault(n => n.Title == "Node 1.2");
            Assert.IsNotNull(n12);
            Assert.AreEqual("/Node1-2.aspx", n12.Url);
            Assert.IsTrue(n12.IsExternal);
            Assert.IsNull(n12.NavigationNode1);

            var n2 = template.Navigation.CurrentNavigation.StructuralNavigation.NavigationNode.FirstOrDefault(n => n.Title == "Node 2");
            Assert.IsNotNull(n2);
            Assert.AreEqual("/Node1.aspx", n2.Url);
            Assert.IsFalse(n2.IsExternal);
            Assert.IsNull(n2.NavigationNode1);

            Assert.AreEqual(false, template.Navigation.SearchNavigation.RemoveExistingNodes);
            Assert.AreEqual("Sample Search Node 01", template.Navigation.SearchNavigation.NavigationNode[0].Title);
            Assert.AreEqual(true, template.Navigation.SearchNavigation.NavigationNode[0].IsExternal);
            Assert.AreEqual(false, template.Navigation.SearchNavigation.NavigationNode[1].IsExternal);
            Assert.AreEqual("{sitecollection}/pages/search01.aspx", template.Navigation.SearchNavigation.NavigationNode[0].Url);
        }

        [TestMethod]
        [TestCategory(TEST_CATEGORY)]
        public void XMLSerializer_Deserialize_SiteFields()
        {
            var provider = new XMLFileSystemTemplateProvider($@"{AppDomain.CurrentDomain.BaseDirectory}\..\..\..\Resources", "Templates");

            var serializer = new TargetSerializer();
            var template = provider.GetTemplate(TEST_TEMPLATE, serializer);

            Assert.IsNotNull(template.SiteFields);
            Assert.AreEqual(4, template.SiteFields.Count);
            Assert.IsNotNull(template.SiteFields.FirstOrDefault(e => e.SchemaXml == "<Field ID=\"{23203E97-3BFE-40CB-AFB4-07AA2B86BF45}\" Type=\"Text\" Name=\"ProjectID\" DisplayName=\"{localization:intranet:ProjectID}\" Group=\"Base.Foundation.Columns\" MaxLength=\"255\" AllowDeletion=\"TRUE\" />"));
            Assert.IsNotNull(template.SiteFields.FirstOrDefault(e => e.SchemaXml == "<Field ID=\"{B01B3DBC-4630-4ED1-B5BA-321BC7841E3D}\" Type=\"Text\" Name=\"ProjectName\" DisplayName=\"{localization:intranet:ProjectName}\" Group=\"Base.Foundation.Columns\" MaxLength=\"255\" AllowDeletion=\"TRUE\" />"));
            Assert.IsNotNull(template.SiteFields.FirstOrDefault(e => e.SchemaXml == "<Field ID=\"{A5DE9600-B7A6-42DD-A05E-10D4F1500208}\" Type=\"Text\" Name=\"ProjectManager\" DisplayName=\"{localization:intranet:ProjectManager}\" Group=\"Base.Foundation.Columns\" MaxLength=\"255\" AllowDeletion=\"TRUE\" />"));
            Assert.IsNotNull(template.SiteFields.FirstOrDefault(e => e.SchemaXml == "<Field ID=\"{F1A1715E-6C52-40DE-8403-E9AAFD0470D0}\" Type=\"Text\" Name=\"DocumentDescription\" DisplayName=\"{localization:core:Description}\" Group=\"Base.Foundation.Columns\" MaxLength=\"255\" AllowDeletion=\"TRUE\" />"));
        }

        [TestMethod]
        [TestCategory(TEST_CATEGORY)]
        public void XMLSerializer_Serialize_SiteFields()
        {
            var provider = new XMLFileSystemTemplateProvider($@"{AppDomain.CurrentDomain.BaseDirectory}\..\..\..\Resources", "Templates");

            var result = new ProvisioningTemplate();
            result.SiteFields.Add(new PnP.Framework.Provisioning.Model.Field()
            {
                SchemaXml = "<Field ID=\"{23203E97-3BFE-40CB-AFB4-07AA2B86BF45}\" Type=\"Text\" Name=\"ProjectID\" DisplayName=\"Project ID\" Group=\"My Columns\" MaxLength=\"255\" AllowDeletion=\"TRUE\" Required=\"TRUE\" />"
            });
            result.SiteFields.Add(new PnP.Framework.Provisioning.Model.Field()
            {
                SchemaXml = "<Field ID = \"{B01B3DBC-4630-4ED1-B5BA-321BC7841E3D}\" Type=\"Text\" Name=\"ProjectName\" DisplayName=\"Project Name\" Group=\"My Columns\" MaxLength=\"255\" AllowDeletion=\"TRUE\" />"
            });
            result.SiteFields.Add(new PnP.Framework.Provisioning.Model.Field()
            {
                SchemaXml = "<Field ID = \"{A5DE9600-B7A6-42DD-A05E-10D4F1500208}\" Type=\"Text\" Name=\"ProjectManager\" DisplayName=\"Project Manager\" Group=\"My Columns\" MaxLength=\"255\" AllowDeletion=\"TRUE\" />"
            });
            result.SiteFields.Add(new PnP.Framework.Provisioning.Model.Field()
            {
                SchemaXml = "<Field ID = \"{F1A1715E-6C52-40DE-8403-E9AAFD0470D0}\" Type=\"Text\" Name=\"DocumentDescription\" DisplayName=\"Document Description\" Group=\"My Columns \" MaxLength=\"255\" AllowDeletion=\"TRUE\" />"
            });

            var serializer = new TargetSerializer();
            provider.SaveAs(result, TEST_OUT_FILE, serializer);

            var path = $"{provider.Connector.Parameters["ConnectionString"]}\\{provider.Connector.Parameters["Container"]}\\{TEST_OUT_FILE}";
            Assert.IsTrue(File.Exists(path)); path = Path.GetFullPath(path);
            var xml = XDocument.Load(path);
            var wrappedResult = XMLSerializer.Deserialize<TargetProvisioning>(xml);

            var template = wrappedResult.Templates[0].ProvisioningTemplate.First();

            Assert.IsNotNull(template.SiteFields);
            Assert.IsNotNull(template.SiteFields.Any);
            Assert.AreEqual(4, template.SiteFields.Any.Length);
            Assert.IsNotNull(template.SiteFields.Any.FirstOrDefault(e => e.OuterXml == "<Field ID=\"{23203E97-3BFE-40CB-AFB4-07AA2B86BF45}\" Type=\"Text\" Name=\"ProjectID\" DisplayName=\"Project ID\" Group=\"My Columns\" MaxLength=\"255\" AllowDeletion=\"TRUE\" Required=\"TRUE\" />"));
            Assert.IsNotNull(template.SiteFields.Any.FirstOrDefault(e => e.OuterXml == "<Field ID=\"{B01B3DBC-4630-4ED1-B5BA-321BC7841E3D}\" Type=\"Text\" Name=\"ProjectName\" DisplayName=\"Project Name\" Group=\"My Columns\" MaxLength=\"255\" AllowDeletion=\"TRUE\" />"));
            Assert.IsNotNull(template.SiteFields.Any.FirstOrDefault(e => e.OuterXml == "<Field ID=\"{A5DE9600-B7A6-42DD-A05E-10D4F1500208}\" Type=\"Text\" Name=\"ProjectManager\" DisplayName=\"Project Manager\" Group=\"My Columns\" MaxLength=\"255\" AllowDeletion=\"TRUE\" />"));
            Assert.IsNotNull(template.SiteFields.Any.FirstOrDefault(e => e.OuterXml == "<Field ID=\"{F1A1715E-6C52-40DE-8403-E9AAFD0470D0}\" Type=\"Text\" Name=\"DocumentDescription\" DisplayName=\"Document Description\" Group=\"My Columns \" MaxLength=\"255\" AllowDeletion=\"TRUE\" />"));
        }

        [TestMethod]
        [TestCategory(TEST_CATEGORY)]
        public void XMLSerializer_Deserialize_ContentTypes()
        {
            var provider = new XMLFileSystemTemplateProvider($@"{AppDomain.CurrentDomain.BaseDirectory}\..\..\..\Resources", "Templates");

            var serializer = new TargetSerializer();
            var template = provider.GetTemplate(TEST_TEMPLATE, serializer);

            Assert.IsNotNull(template.ContentTypes);

            var ct = template.ContentTypes.FirstOrDefault(c => c.Id == "0x01005D4F34E4BE7F4B6892AEBE088EDD215E");
            Assert.IsNotNull(ct);
            Assert.AreEqual("0x01005D4F34E4BE7F4B6892AEBE088EDD215E", ct.Id);
            Assert.AreEqual("General Project Document", ct.Name);
            Assert.AreEqual("General Project Document Content Type", ct.Description);
            Assert.AreEqual("Base Foundation Content Types", ct.Group);
            Assert.AreEqual("/Forms/DisplayForm.aspx", ct.DisplayFormUrl);
            Assert.AreEqual("/Forms/NewForm.aspx", ct.NewFormUrl);
            Assert.AreEqual("/Forms/EditForm.aspx", ct.EditFormUrl);
            Assert.AreEqual("DocumentTemplate.dotx", ct.DocumentTemplate);
            Assert.AreEqual(new Guid("F1A1715E-6C52-40DE-8403-E9AAFD0470D0"), ct.FieldRefs[3].Id);
            Assert.AreEqual(true, ct.FieldRefs[3].UpdateChildren);
            Assert.IsFalse(ct.Hidden);
            Assert.IsFalse(ct.Overwrite);
            Assert.IsFalse(ct.ReadOnly);
            Assert.IsFalse(ct.Sealed);
            Assert.IsTrue(ct.UpdateChildren);
            Assert.AreEqual("49b0ccdd-1bdf-47ba-a92f-dfddec265d63", ct.DisplayFormClientSideComponentId);
            Assert.AreEqual("{'sampleProperty':'Sample value'}", ct.DisplayFormClientSideComponentProperties);
            Assert.AreEqual("7c7c9c44-4dd4-4a2b-8619-beba16e08efa", ct.NewFormClientSideComponentId);
            Assert.AreEqual("{'sampleProperty':'Sample value'}", ct.NewFormClientSideComponentProperties);
            Assert.AreEqual("3c4825e4-4b20-496d-9450-cd76f5135323", ct.EditFormClientSideComponentId);
            Assert.AreEqual("{'sampleProperty':'Sample value'}", ct.EditFormClientSideComponentProperties);

            ct = template.ContentTypes.FirstOrDefault(c => c.Id == "0x0120D5200039D83CD2C9BA4A4499AEE6BE3562E023");
            Assert.IsNotNull(ct.DocumentSetTemplate);
            Assert.IsNotNull(ct.DocumentSetTemplate.AllowedContentTypes);
            Assert.AreEqual("{sitecollection}/_cts/ProjectDocumentSet/ProjectHomePage.aspx", ct.DocumentSetTemplate.WelcomePage);
            Assert.IsTrue(ct.DocumentSetTemplate.RemoveExistingContentTypes);
            Assert.IsFalse(ct.DocumentSetTemplate.UpdateChildren);

            var act = ct.DocumentSetTemplate.AllowedContentTypes.FirstOrDefault(c => c.ContentTypeId == "0x01005D4F34E4BE7F4B6892AEBE088EDD215E");
            Assert.IsNotNull(act);
            Assert.AreEqual(act.ContentTypeId, "0x01005D4F34E4BE7F4B6892AEBE088EDD215E");
            Assert.AreEqual(act.Name, "General Project Document");
            Assert.IsFalse(act.Remove);

            var sf = ct.DocumentSetTemplate.SharedFields.FirstOrDefault(c => c.Id == new Guid("23203E97-3BFE-40CB-AFB4-07AA2B86BF45"));
            Assert.IsNotNull(sf);
            Assert.AreEqual(sf.Id, new Guid("23203E97-3BFE-40CB-AFB4-07AA2B86BF45"));
            Assert.AreEqual(sf.Name, "ProjectID");
            Assert.IsFalse(sf.Remove);

            var wpf = ct.DocumentSetTemplate.WelcomePageFields.FirstOrDefault(c => c.Id == new Guid("23203E97-3BFE-40CB-AFB4-07AA2B86BF45"));
            Assert.IsNotNull(wpf);
            Assert.AreEqual(wpf.Id, new Guid("23203E97-3BFE-40CB-AFB4-07AA2B86BF45"));
            Assert.AreEqual(wpf.Name, "ProjectID");
            Assert.IsFalse(wpf.Remove);

            Assert.IsNotNull(ct.DocumentSetTemplate.DefaultDocuments);

            var defaultDocument = ct.DocumentSetTemplate.DefaultDocuments.FirstOrDefault(d => d.ContentTypeId == "0x01005D4F34E4BE7F4B6892AEBE088EDD215E");
            Assert.IsNotNull(defaultDocument);
            Assert.AreEqual("./ProjectDocumentSet/ProjectMainDocument.docx", defaultDocument.FileSourcePath);
            Assert.AreEqual("ProjectMainDocument.docx", defaultDocument.Name);

            var xmlDocs = ct.DocumentSetTemplate.XmlDocuments;
            Assert.IsNotNull(xmlDocs);
            Assert.AreEqual(4, xmlDocs.Elements("XmlDocument").Count());
        }

        [TestMethod]
        [TestCategory(TEST_CATEGORY)]
        public void XMLSerializer_Serialize_ContentTypes()
        {
            var provider = new XMLFileSystemTemplateProvider($@"{AppDomain.CurrentDomain.BaseDirectory}\..\..\..\Resources", "Templates");

            var result = new ProvisioningTemplate();

            var contentType = new PnP.Framework.Provisioning.Model.ContentType()
            {
                Id = "0x01005D4F34E4BE7F4B6892AEBE088EDD215E",
                Name = "General Project Document",
                Description = "General Project Document Content Type",
                Group = "Base Foundation Content Types",
                DisplayFormUrl = "/Forms/DisplayForm.aspx",
                NewFormUrl = "/Forms/NewForm.aspx",
                EditFormUrl = "/Forms/EditForm.aspx",
                DocumentTemplate = "DocumentTemplate.dotx",
                Hidden = true,
                Overwrite = true,
                ReadOnly = true,
                Sealed = true,
                UpdateChildren = true,
                DisplayFormClientSideComponentId = "49b0ccdd-1bdf-47ba-a92f-dfddec265d63",
                DisplayFormClientSideComponentProperties = "{'sampleProperty':'Sample value'}",
                NewFormClientSideComponentId = "7c7c9c44-4dd4-4a2b-8619-beba16e08efa",
                NewFormClientSideComponentProperties = "{'sampleProperty':'Sample value'}",
                EditFormClientSideComponentId = "3c4825e4-4b20-496d-9450-cd76f5135323",
                EditFormClientSideComponentProperties = "{'sampleProperty':'Sample value'}"
            };

            var documentSetTemplate = new DocumentSetTemplate { RemoveExistingContentTypes = true, UpdateChildren = true };
            documentSetTemplate.AllowedContentTypes.Add(new ContentTypeReference { ContentTypeId = "0x01005D4F34E4BE7F4B6892AEBE088EDD215E002", Name = "TestCT", Remove = true });
            documentSetTemplate.SharedFields.Add(new FieldReference { Id = new Guid("f6e7bdd5-bdcb-4c72-9f18-2bd8c27003d3"), Name = "Field1", Remove = false });
            documentSetTemplate.SharedFields.Add(new FieldReference { Id = new Guid("a8df65ec-0d06-4df1-8edf-55d48b3936dc"), Name = "Field2", Remove = true });
            documentSetTemplate.WelcomePageFields.Add(new FieldReference { Id = new Guid("f6e7bdd5-bdcb-4c72-9f18-2bd8c27003d3"), Name = "Field1", Remove = false });
            documentSetTemplate.WelcomePageFields.Add(new FieldReference { Id = new Guid("a8df65ec-0d06-4df1-8edf-55d48b3936dc"), Name = "Field2", Remove = true });
            documentSetTemplate.WelcomePage = "home.aspx";
            documentSetTemplate.DefaultDocuments.Add(new DefaultDocument()
            {
                ContentTypeId = "0x01005D4F34E4BE7F4B6892AEBE088EDD215E001",
                FileSourcePath = "document.dotx",
                Name = "DefaultDocument"
            });
            contentType.DocumentSetTemplate = documentSetTemplate;
            contentType.FieldRefs.Add(new FieldRef("TestField")
            {
                Id = new Guid("23203e97-3bfe-40cb-afb4-07aa2b86bf45"),
                Required = true,
                Hidden = true
            });
            contentType.FieldRefs.Add(new FieldRef("TestField1"));
            contentType.FieldRefs.Add(new FieldRef("TestField2"));
            contentType.FieldRefs.Add(new FieldRef("TestField3"));
            result.ContentTypes.Add(contentType);

            var serializer = new TargetSerializer();
            provider.SaveAs(result, TEST_OUT_FILE, serializer);

            var path = $"{provider.Connector.Parameters["ConnectionString"]}\\{provider.Connector.Parameters["Container"]}\\{TEST_OUT_FILE}";
            Assert.IsTrue(File.Exists(path)); path = Path.GetFullPath(path);
            var xml = XDocument.Load(path);
            var wrappedResult = XMLSerializer.Deserialize<TargetProvisioning>(xml);

            var template = wrappedResult.Templates[0].ProvisioningTemplate.First();

            Assert.IsNotNull(template.ContentTypes);

            var ct = template.ContentTypes.FirstOrDefault();
            Assert.IsNotNull(ct);

            Assert.AreEqual("0x01005D4F34E4BE7F4B6892AEBE088EDD215E", ct.ID);
            Assert.AreEqual("General Project Document", ct.Name);
            Assert.AreEqual("General Project Document Content Type", ct.Description);
            Assert.AreEqual("Base Foundation Content Types", ct.Group);
            Assert.AreEqual("/Forms/DisplayForm.aspx", ct.DisplayFormUrl);
            Assert.AreEqual("/Forms/NewForm.aspx", ct.NewFormUrl);
            Assert.AreEqual("/Forms/EditForm.aspx", ct.EditFormUrl);
            Assert.IsNotNull(ct.DocumentTemplate);
            Assert.AreEqual("DocumentTemplate.dotx", ct.DocumentTemplate.TargetName);
            Assert.IsTrue(ct.Hidden);
            Assert.IsTrue(ct.Overwrite);
            Assert.IsTrue(ct.ReadOnly);
            Assert.IsTrue(ct.Sealed);
            Assert.IsTrue(ct.UpdateChildren);
            Assert.AreEqual("49b0ccdd-1bdf-47ba-a92f-dfddec265d63", ct.DisplayFormClientSideComponentId);
            Assert.AreEqual("{'sampleProperty':'Sample value'}", ct.DisplayFormClientSideComponentProperties);
            Assert.AreEqual("7c7c9c44-4dd4-4a2b-8619-beba16e08efa", ct.NewFormClientSideComponentId);
            Assert.AreEqual("{'sampleProperty':'Sample value'}", ct.NewFormClientSideComponentProperties);
            Assert.AreEqual("3c4825e4-4b20-496d-9450-cd76f5135323", ct.EditFormClientSideComponentId);
            Assert.AreEqual("{'sampleProperty':'Sample value'}", ct.EditFormClientSideComponentProperties);

            Assert.IsNotNull(ct.DocumentSetTemplate);
            Assert.IsTrue(ct.DocumentSetTemplate.UpdateChildren);
            Assert.IsNotNull(ct.DocumentSetTemplate.AllowedContentTypes);

            var act = ct.DocumentSetTemplate.AllowedContentTypes.AllowedContentType.FirstOrDefault();
            Assert.IsNotNull(act);
            Assert.AreEqual(act.ContentTypeID, "0x01005D4F34E4BE7F4B6892AEBE088EDD215E002");
            Assert.AreEqual(act.Name, "TestCT");
            Assert.IsTrue(act.Remove);
            Assert.AreEqual(true, ct.DocumentSetTemplate.AllowedContentTypes.RemoveExistingContentTypes);
            Assert.IsNotNull(ct.DocumentSetTemplate.SharedFields);

            var sf1 = ct.DocumentSetTemplate.SharedFields.FirstOrDefault(c => c.ID == "f6e7bdd5-bdcb-4c72-9f18-2bd8c27003d3");
            Assert.IsNotNull(sf1);
            Assert.AreEqual(sf1.ID, "f6e7bdd5-bdcb-4c72-9f18-2bd8c27003d3");
            Assert.AreEqual(sf1.Name, "Field1");
            Assert.IsFalse(sf1.Remove);

            var sf2 = ct.DocumentSetTemplate.SharedFields.FirstOrDefault(c => c.ID == "a8df65ec-0d06-4df1-8edf-55d48b3936dc");
            Assert.IsNotNull(sf2);
            Assert.AreEqual(sf2.ID, "a8df65ec-0d06-4df1-8edf-55d48b3936dc");
            Assert.AreEqual(sf2.Name, "Field2");
            Assert.IsTrue(sf2.Remove);

            var wpf1 = ct.DocumentSetTemplate.SharedFields.FirstOrDefault(c => c.ID == "f6e7bdd5-bdcb-4c72-9f18-2bd8c27003d3");
            Assert.IsNotNull(wpf1);
            Assert.AreEqual(wpf1.ID, "f6e7bdd5-bdcb-4c72-9f18-2bd8c27003d3");
            Assert.AreEqual(wpf1.Name, "Field1");
            Assert.IsFalse(wpf1.Remove);

            var wpf2 = ct.DocumentSetTemplate.SharedFields.FirstOrDefault(c => c.ID == "a8df65ec-0d06-4df1-8edf-55d48b3936dc");
            Assert.IsNotNull(wpf2);
            Assert.AreEqual(wpf2.ID, "a8df65ec-0d06-4df1-8edf-55d48b3936dc");
            Assert.AreEqual(wpf2.Name, "Field2");
            Assert.IsTrue(wpf2.Remove);

            Assert.AreEqual("home.aspx", ct.DocumentSetTemplate.WelcomePage);
            Assert.IsNotNull(ct.DocumentSetTemplate.DefaultDocuments);

            var dd = ct.DocumentSetTemplate.DefaultDocuments.FirstOrDefault(d => d.ContentTypeID == "0x01005D4F34E4BE7F4B6892AEBE088EDD215E001");
            Assert.IsNotNull(dd);
            Assert.AreEqual("document.dotx", dd.FileSourcePath);
            Assert.AreEqual("DefaultDocument", dd.Name);

            Assert.IsNotNull(ct.FieldRefs);
            Assert.AreEqual(4, ct.FieldRefs.Count());

            var field = ct.FieldRefs.FirstOrDefault(f => f.Name == "TestField");
            Assert.IsNotNull(field);
            Assert.AreEqual("23203e97-3bfe-40cb-afb4-07aa2b86bf45", field.ID);
            Assert.IsTrue(field.Required);
            Assert.IsTrue(field.Hidden);
        }

        [TestMethod]
        [TestCategory(TEST_CATEGORY)]
        public void XMLSerializer_Deserialize_ListInstances()
        {
            var provider = new XMLFileSystemTemplateProvider($@"{AppDomain.CurrentDomain.BaseDirectory}\..\..\..\Resources", "Templates");

            var serializer = new TargetSerializer();
            var template = provider.GetTemplate(TEST_TEMPLATE, serializer);

            Assert.IsNotNull(template.Lists);
            Assert.AreEqual(3, template.Lists.Count);

            // common properties
            var list = template.Lists.FirstOrDefault(ls => ls.Title == "{parameter:CompanyName} - Projects");
            Assert.IsNotNull(list);
            Assert.AreEqual("Project Documents are stored here", list.Description);
            Assert.AreEqual("", list.DocumentTemplate);
            Assert.IsTrue(list.OnQuickLaunch);
            Assert.AreEqual(101, list.TemplateType);
            Assert.AreEqual("Lists/Projects", list.Url);
            Assert.IsTrue(list.EnableVersioning);
            Assert.IsTrue(list.ForceCheckout);
            Assert.IsFalse(list.RemoveExistingContentTypes);
            Assert.IsTrue(list.NoCrawl);
            Assert.AreEqual(PnP.Framework.Provisioning.Model.ListExperience.ClassicExperience, list.ListExperience);
            Assert.IsTrue(list.ContentTypesEnabled);
            Assert.AreEqual("./Forms/Display.aspx", list.DefaultDisplayFormUrl);
            Assert.AreEqual("./Forms/Edit.aspx", list.DefaultEditFormUrl);
            Assert.AreEqual("./Forms/NewItem.aspx", list.DefaultNewFormUrl);
            Assert.AreEqual(ListReadingDirection.LTR, list.Direction);
            Assert.AreEqual(1, list.DraftVersionVisibility);
            Assert.IsFalse(list.EnableAttachments);
            Assert.IsTrue(list.EnableFolderCreation);
            Assert.IsTrue(list.EnableMinorVersions);
            Assert.IsFalse(list.EnableModeration);
            Assert.IsFalse(list.Hidden);
            Assert.AreEqual("./Images/Sample.png", list.ImageUrl);
            Assert.AreEqual(true, list.IrmExpire);
            Assert.AreEqual(false, list.IrmReject);
            Assert.AreEqual(false, list.IsApplicationList);
            Assert.AreEqual(10, list.MaxVersionLimit);
            Assert.AreEqual(10, list.MinorVersionLimit);
            Assert.AreEqual(11, list.ReadSecurity);
            Assert.AreEqual(new Guid("81a7b6a8-c0e9-4819-aea1-8fc8894d3c43"), list.TemplateFeatureID);
            Assert.AreEqual("sample formula here", list.ValidationFormula);
            Assert.AreEqual("validation message here", list.ValidationMessage);
            Assert.AreEqual("fake-template.stp", list.TemplateInternalName);
            Assert.IsTrue(list.EnableAudienceTargeting);
            Assert.IsTrue(list.EnableClassicAudienceTargeting);
            Assert.AreEqual(120, list.Webhooks[0].ExpiresInDays);
            Assert.AreEqual("http://myapp.azurewebsites.net/WebHookListener", list.Webhooks[0].ServerNotificationUrl);

            // security
            var security = list.Security;
            Assert.IsNotNull(security);
            Assert.IsTrue(security.ClearSubscopes);
            Assert.IsTrue(security.CopyRoleAssignments);
            Assert.IsNotNull(security.RoleAssignments);
            Assert.AreEqual(2, security.RoleAssignments.Count);
            var roleAssignment = security.RoleAssignments.FirstOrDefault(r => r.Principal == "Power Users");
            Assert.IsNotNull(roleAssignment);
            Assert.AreEqual("Full Control", roleAssignment.RoleDefinition);
            roleAssignment = security.RoleAssignments.FirstOrDefault(r => r.Principal == "Guests");
            Assert.IsNotNull(roleAssignment);
            Assert.AreEqual("View Only", roleAssignment.RoleDefinition);

            Assert.IsNotNull(list.ContentTypeBindings);
            Assert.AreEqual(2, list.ContentTypeBindings.Count);
            var ct = list.ContentTypeBindings.FirstOrDefault(c => c.ContentTypeId == "0x01005D4F34E4BE7F4B6892AEBE088EDD215E");
            Assert.IsNotNull(ct);
            Assert.IsTrue(ct.Default);
            Assert.IsFalse(ct.Remove);
            ct = list.ContentTypeBindings.FirstOrDefault(c => c.ContentTypeId == "0x0120");
            Assert.IsNotNull(ct);
            Assert.IsFalse(ct.Default);
            Assert.IsTrue(ct.Remove);

            Assert.IsNotNull(list.FieldDefaults);
            Assert.AreEqual(3, list.FieldDefaults.Count);
            var fd = list.FieldDefaults.FirstOrDefault(f => f.Key == "ProjectName");
            Assert.IsNotNull(fd);
            Assert.AreEqual("Default Project Name", fd.Value);
            fd = list.FieldDefaults.FirstOrDefault(f => f.Key == "ProjectManager");
            Assert.IsNotNull(fd);
            Assert.AreEqual("[Me]", fd.Value);
            fd = list.FieldDefaults.FirstOrDefault(f => f.Key == "ProjectDescription");
            Assert.IsNotNull(fd);
            Assert.AreEqual("Default Project Description", fd.Value);

            // Default Column Values
            Assert.IsNotNull(list.DefaultColumnValues);
            Assert.AreEqual(2, list.DefaultColumnValues.Count);
            Assert.AreEqual("Custom Value 1", list.DefaultColumnValues["Field1"]);
            Assert.AreEqual("Custom Value 2", list.DefaultColumnValues["Field2"]);

            Assert.IsNotNull(list.DataRows);
            Assert.AreEqual(2, list.DataRows.Count);
            Assert.AreEqual("ProjectID", list.DataRows.KeyColumn);
            Assert.AreEqual(UpdateBehavior.Overwrite, list.DataRows.UpdateBehavior);

            #region data row 1 asserts

            var dataRow = list.DataRows.FirstOrDefault(r => r.Values.Any(d => d.Value.StartsWith("PRJ01")));
            Assert.IsNotNull(dataRow);

            security = dataRow.Security;
            Assert.IsNotNull(security);
            Assert.IsTrue(security.ClearSubscopes);
            Assert.IsTrue(security.CopyRoleAssignments);
            Assert.IsNotNull(security.RoleAssignments);
            Assert.AreEqual(3, security.RoleAssignments.Count);
            roleAssignment = security.RoleAssignments.FirstOrDefault(r => r.Principal == "user1@contoso.com");
            Assert.IsNotNull(roleAssignment);
            Assert.AreEqual("Full Control", roleAssignment.RoleDefinition);
            roleAssignment = security.RoleAssignments.FirstOrDefault(r => r.Principal == "user2@contoso.com");
            Assert.IsNotNull(roleAssignment);
            Assert.AreEqual("Edit", roleAssignment.RoleDefinition);
            roleAssignment = security.RoleAssignments.FirstOrDefault(r => r.Principal == "user3@contoso.com");
            Assert.IsNotNull(roleAssignment);
            Assert.AreEqual("View Only", roleAssignment.RoleDefinition);

            var dataValues = dataRow.Values.FirstOrDefault(d => d.Key == "ProjectID");
            Assert.IsNotNull(dataValues);
            Assert.AreEqual("PRJ01", dataValues.Value);
            dataValues = dataRow.Values.FirstOrDefault(d => d.Key == "ProjectName");
            Assert.IsNotNull(dataValues);
            Assert.AreEqual("Sample Project 01", dataValues.Value);
            dataValues = dataRow.Values.FirstOrDefault(d => d.Key == "ProjectManager");
            Assert.IsNotNull(dataValues);
            Assert.AreEqual("Me", dataValues.Value);
            dataValues = dataRow.Values.FirstOrDefault(d => d.Key == "ProjectDescription");
            Assert.IsNotNull(dataValues);
            Assert.AreEqual("This is a sample Project", dataValues.Value);

            var attachments = dataRow.Attachments;
            Assert.IsNotNull(attachments);
            Assert.AreEqual(2, attachments.Count);
            Assert.AreEqual("OneAttachment.docx", attachments[0].Name);
            Assert.AreEqual("./Attachments/OneAttachment.docx", attachments[0].Src);
            Assert.AreEqual(false, attachments[0].Overwrite);
            Assert.AreEqual("AnotherAttachment.pptx", attachments[1].Name);
            Assert.AreEqual("./Attachments/AnotherAttachment.pptx", attachments[1].Src);
            Assert.AreEqual(true, attachments[1].Overwrite);

            #endregion

            #region data row 2 asserts

            dataRow = list.DataRows.FirstOrDefault(r => r.Values.Any(d => d.Value.StartsWith("PRJ021")));
            Assert.IsNotNull(dataRow);

            security = dataRow.Security;
            Assert.IsNotNull(security);
            Assert.IsTrue(security.ClearSubscopes);
            Assert.IsFalse(security.CopyRoleAssignments);
            Assert.IsNotNull(security.RoleAssignments);
            Assert.AreEqual(3, security.RoleAssignments.Count);
            roleAssignment = security.RoleAssignments.FirstOrDefault(r => r.Principal == "user1@contoso.com");
            Assert.IsNotNull(roleAssignment);
            Assert.AreEqual("View Only", roleAssignment.RoleDefinition);
            roleAssignment = security.RoleAssignments.FirstOrDefault(r => r.Principal == "user2@contoso.com");
            Assert.IsNotNull(roleAssignment);
            Assert.AreEqual("Edit", roleAssignment.RoleDefinition);
            roleAssignment = security.RoleAssignments.FirstOrDefault(r => r.Principal == "user3@contoso.com");
            Assert.IsNotNull(roleAssignment);
            Assert.AreEqual("Full Control", roleAssignment.RoleDefinition);

            dataValues = dataRow.Values.FirstOrDefault(d => d.Key == "ProjectID");
            Assert.IsNotNull(dataValues);
            Assert.AreEqual("PRJ021", dataValues.Value);
            dataValues = dataRow.Values.FirstOrDefault(d => d.Key == "ProjectName");
            Assert.IsNotNull(dataValues);
            Assert.AreEqual("Sample Project 02", dataValues.Value);
            dataValues = dataRow.Values.FirstOrDefault(d => d.Key == "ProjectManager");
            Assert.IsNotNull(dataValues);
            Assert.AreEqual("You", dataValues.Value);
            dataValues = dataRow.Values.FirstOrDefault(d => d.Key == "ProjectDescription");
            Assert.IsNotNull(dataValues);
            Assert.AreEqual("This is another sample Project", dataValues.Value);

            #endregion

            #region user custom action
            Assert.IsNotNull(list.UserCustomActions);
            Assert.AreEqual(1, list.UserCustomActions.Count);
            var ua = list.UserCustomActions.FirstOrDefault(a => a.Name == "CA_LIST_ECB_ITEM");
            Assert.IsNotNull(ua);
            Assert.AreEqual("Custom Action for ECB in List", ua.Description);
            Assert.IsTrue(ua.Enabled);
            Assert.AreEqual("EditControlBlock", ua.Location);
            Assert.AreEqual("0x01005D4F34E4BE7F4B6892AEBE088EDD215E", ua.RegistrationId);
            Assert.AreEqual(UserCustomActionRegistrationType.ContentType, ua.RegistrationType);
            Assert.AreEqual(1000, ua.Sequence);
            Assert.AreEqual("https://spmanaged.azurewebsites.net/customAction/?ItemUrl={ItemUrl}&ItemId={ItemId}&ListId={ListId}&SiteUrl={SiteUrl}", ua.Url);
            Assert.AreEqual("Sample Custom Action", ua.Title);
            #endregion

            #region folders

            Assert.IsNotNull(list.Folders);
            Assert.AreEqual(4, list.Folders.Count);
            var fl = list.Folders.FirstOrDefault(f => f.Name == "SubFolder-01");
            Assert.IsNotNull(fl);
            Assert.IsTrue(fl.Folders.Count == 1);
            fl = list.Folders.FirstOrDefault(f => f.Name == "SubFolder-02");
            Assert.IsNotNull(fl);
            Assert.IsNotNull(fl.Folders);
            var fl1 = fl.Folders.FirstOrDefault(f => f.Name == "SubFolder-02-01");
            Assert.IsNotNull(fl1);
            Assert.IsTrue(fl1.Folders.Count == 1);
            fl1 = fl1.Folders.FirstOrDefault(f => f.Name == "SubFolder-02-01-01");
            Assert.IsTrue(fl1.Folders == null || fl1.Folders.Count == 0);
            security = fl1.Security;
            Assert.IsNotNull(security);
            Assert.IsTrue(security.ClearSubscopes);
            Assert.IsFalse(security.CopyRoleAssignments);
            Assert.IsNotNull(security.RoleAssignments);
            Assert.AreEqual(3, security.RoleAssignments.Count);
            roleAssignment = security.RoleAssignments.FirstOrDefault(r => r.Principal == "user3@contoso.com");
            Assert.IsNotNull(roleAssignment);
            Assert.AreEqual("Full Control", roleAssignment.RoleDefinition);
            roleAssignment = security.RoleAssignments.FirstOrDefault(r => r.Principal == "user2@contoso.com");
            Assert.IsNotNull(roleAssignment);
            Assert.AreEqual("Edit", roleAssignment.RoleDefinition);
            roleAssignment = security.RoleAssignments.FirstOrDefault(r => r.Principal == "user1@contoso.com");
            Assert.IsNotNull(roleAssignment);
            Assert.AreEqual("View Only", roleAssignment.RoleDefinition);

            var flDocumentSet = list.Folders.FirstOrDefault(f => f.Name == "Sample-DocumentSet");
            Assert.AreEqual("0x0120001234567890", flDocumentSet.ContentTypeID);
            Assert.IsNotNull(flDocumentSet.Properties);
            Assert.AreEqual(2, flDocumentSet.Properties.Count);
            Assert.AreEqual("CustomValue01", flDocumentSet.Properties["CustomProperty01"]);
            Assert.AreEqual("CustomValue02", flDocumentSet.Properties["CustomProperty02"]);

            #endregion

            #region IRM Settings
            Assert.AreEqual(true, list.IRMSettings.AllowPrint);
            Assert.AreEqual(false, list.IRMSettings.AllowWriteCopy);
            Assert.AreEqual(true, list.IRMSettings.AllowScript);
            Assert.AreEqual(true, list.IRMSettings.DisableDocumentBrowserView);
            Assert.AreEqual(30, list.IRMSettings.DocumentAccessExpireDays);
            Assert.AreEqual(90, list.IRMSettings.DocumentLibraryProtectionExpiresInDays);
            Assert.AreEqual(true, list.IRMSettings.Enabled);
            Assert.AreEqual(true, list.IRMSettings.EnableDocumentAccessExpire);
            Assert.AreEqual(true, list.IRMSettings.EnableDocumentBrowserPublishingView);
            Assert.AreEqual(false, list.IRMSettings.EnableGroupProtection);
            Assert.AreEqual(false, list.IRMSettings.EnableLicenseCacheExpire);
            Assert.AreEqual("Custom IRM Group", list.IRMSettings.GroupName);
            Assert.AreEqual(60, list.IRMSettings.LicenseCacheExpireDays);
            Assert.AreEqual("Sample IRM Policy", list.IRMSettings.PolicyDescription);
            Assert.AreEqual("Sample IRM Policy", list.IRMSettings.PolicyTitle);
            #endregion
        }

        [TestMethod]
        [TestCategory(TEST_CATEGORY)]
        public void XMLSerializer_Serialize_ListInstances()
        {
            var provider = new XMLFileSystemTemplateProvider($@"{AppDomain.CurrentDomain.BaseDirectory}\..\..\..\Resources", "Templates");

            var result = new ProvisioningTemplate();
            var list = new PnP.Framework.Provisioning.Model.ListInstance()
            {
                Title = "Project Documents",
                Description = "Project Documents are stored here",
                DocumentTemplate = "document.dotx",
                OnQuickLaunch = true,
                TemplateType = 101,
                Url = "/Lists/ProjectDocuments",
                EnableVersioning = true,
                ForceCheckout = true,
                RemoveExistingContentTypes = true,
                NoCrawl = true,
                ListExperience = Provisioning.Model.ListExperience.ClassicExperience,
                ContentTypesEnabled = true,
                DefaultDisplayFormUrl = "./Forms/Display.aspx",
                DefaultEditFormUrl = "./Forms/Edit.aspx",
                DefaultNewFormUrl = "./Forms/NewItem.aspx",
                Direction = ListReadingDirection.LTR,
                DraftVersionVisibility = 1,
                EnableAttachments = true,
                EnableFolderCreation = true,
                EnableMinorVersions = true,
                EnableModeration = true,
                Hidden = true,
                ImageUrl = "./Images/Sample.png",
                IrmExpire = true,
                IrmReject = false,
                IsApplicationList = false,
                MaxVersionLimit = 10,
                MinorVersionLimit = 2,
                ReadSecurity = 11,
                TemplateFeatureID = new Guid("30FB193E-016E-45A6-B6FD-C6C2B31AA150"),
                ValidationFormula = "sample formula here",
                ValidationMessage = "validation message here",
                TemplateInternalName = "fake-template.stp",
                EnableAudienceTargeting = true,
                EnableClassicAudienceTargeting = true,
                RemoveExistingViews = true,
                Security = new PnP.Framework.Provisioning.Model.ObjectSecurity(new List<PnP.Framework.Provisioning.Model.RoleAssignment>()
                {
                    new PnP.Framework.Provisioning.Model.RoleAssignment()
                    {
                        Principal="Principal01",
                        RoleDefinition ="Read"
                    },
                    new PnP.Framework.Provisioning.Model.RoleAssignment()
                    {
                        Principal="Principal02",
                        RoleDefinition ="Contribute"
                    },
                    new PnP.Framework.Provisioning.Model.RoleAssignment()
                    {
                        Principal="Principal03",
                        RoleDefinition ="FullControl"
                    }
                })
                {
                    ClearSubscopes = true,
                    CopyRoleAssignments = true,

                }
            };
            list.ContentTypeBindings.Add(new PnP.Framework.Provisioning.Model.ContentTypeBinding()
            {
                ContentTypeId = "0x01005D4F34E4BE7F4B6892AEBE088EDD215E",
                Default = true
            });
            list.ContentTypeBindings.Add(new PnP.Framework.Provisioning.Model.ContentTypeBinding()
            {
                ContentTypeId = "0x0101",
                Remove = true
            });
            list.ContentTypeBindings.Add(new PnP.Framework.Provisioning.Model.ContentTypeBinding()
            {
                ContentTypeId = "0x0102"
            });

            list.DefaultColumnValues.Add("Field1", "Custom Value 1");
            list.DefaultColumnValues.Add("Field2", "Custom Value 2");

            list.FieldDefaults.Add("Field01", "DefaultValue01");
            list.FieldDefaults.Add("Field02", "DefaultValue02");
            list.FieldDefaults.Add("Field03", "DefaultValue03");
            list.FieldDefaults.Add("Field04", "DefaultValue04");

            list.Webhooks.Add(new PnP.Framework.Provisioning.Model.Webhook
            {
                ExpiresInDays = 120,
                ServerNotificationUrl = "http://myapp.azurewebsites.net/WebHookListener"
            });

            list.IRMSettings = new PnP.Framework.Provisioning.Model.IRMSettings
            {
                AllowPrint = true,
                AllowWriteCopy = false,
                AllowScript = true,
                DisableDocumentBrowserView = true,
                DocumentAccessExpireDays = 30,
                DocumentLibraryProtectionExpiresInDays = 90,
                Enabled = true,
                EnableDocumentAccessExpire = true,
                EnableDocumentBrowserPublishingView = true,
                EnableGroupProtection = false,
                EnableLicenseCacheExpire = false,
                GroupName = "Custom IRM Group",
                LicenseCacheExpireDays = 60,
                PolicyDescription = "Sample IRM Policy",
                PolicyTitle = "Sample IRM Policy"
            };

            #region data rows
            list.DataRows.Add(new DataRow(new Dictionary<string, string>() {
                { "Field01", "Value01-01" },
                { "Field02", "Value01-02" },
                { "Field03", "Value01-03" },
                { "Field04", "Value01-04" },
            },
            new PnP.Framework.Provisioning.Model.ObjectSecurity(new List<PnP.Framework.Provisioning.Model.RoleAssignment>() {
                new PnP.Framework.Provisioning.Model.RoleAssignment()
                {
                    Principal ="Principal01",
                    RoleDefinition ="Read"
                },
                new PnP.Framework.Provisioning.Model.RoleAssignment()
                {
                    Principal ="Principal02",
                    RoleDefinition ="Contribute"
                }
                ,
                new PnP.Framework.Provisioning.Model.RoleAssignment()
                {
                    Principal ="Principal03",
                    RoleDefinition ="FullControl"
                }
            })
            {
                ClearSubscopes = true,
                CopyRoleAssignments = true
            }));
            list.DataRows.Add(new DataRow(new Dictionary<string, string>() {
                { "Field01", "Value02-01" },
                { "Field02", "Value02-02" },
                { "Field03", "Value02-03" },
                { "Field04", "Value02-04" },
            },
            new PnP.Framework.Provisioning.Model.ObjectSecurity(new List<PnP.Framework.Provisioning.Model.RoleAssignment>() {
                new PnP.Framework.Provisioning.Model.RoleAssignment()
                {
                    Principal ="Principal01",
                    RoleDefinition ="Read"
                },
                new PnP.Framework.Provisioning.Model.RoleAssignment()
                {
                    Principal ="Principal02",
                    RoleDefinition ="Contribute"
                }
                ,
                new PnP.Framework.Provisioning.Model.RoleAssignment()
                {
                    Principal ="Principal03",
                    RoleDefinition ="FullControl"
                }
            })
            {
                ClearSubscopes = false,
                CopyRoleAssignments = false
            }));
            list.DataRows.Add(new DataRow(new Dictionary<string, string>() {
                { "Field01", "Value03-01" },
                { "Field02", "Value03-02" },
                { "Field03", "Value03-03" },
                { "Field04", "Value03-04" },
            }));
            #endregion

            var ca = new PnP.Framework.Provisioning.Model.CustomAction()
            {
                Name = "SampleCustomAction",
                Description = "Just a sample custom action",
                Enabled = true,
                Group = "Samples",
                ImageUrl = "OneImage.png",
                Location = "Any",
                RegistrationId = "0x0101",
                RegistrationType = UserCustomActionRegistrationType.ContentType,
                Sequence = 100,
                ScriptBlock = "scriptblock",
                ScriptSrc = "script.js",
                Url = "http://somewhere.com/",
                Rights = new BasePermissions(),
                Title = "Sample Action",
                Remove = true,
                CommandUIExtension = XElement.Parse("<CommandUIExtension><customElement><!--Whateveryoulikehere--></customElement></CommandUIExtension>")
            };
            ca.Rights.Set(PermissionKind.AddListItems);
            list.UserCustomActions.Add(ca);

            #region views
            list.Views.Add(new PnP.Framework.Provisioning.Model.View()
            {
                SchemaXml = @"<View DisplayName=""View One"">
                  <ViewFields>
                    <FieldRef Name=""ID"" />
                    <FieldRef Name=""Title"" />
                    <FieldRef Name=""ProjectID"" />
                    <FieldRef Name=""ProjectName"" />
                    <FieldRef Name=""ProjectManager"" />
                    <FieldRef Name=""DocumentDescription"" />
                  </ViewFields>
                  <Query>
                    <Where>
                      <Eq>
                        <FieldRef Name=""ProjectManager"" />
                        <Value Type=""Text"">[Me]</Value>
                      </Eq>
                    </Where>
                  </Query>
                </View>"
            });
            list.Views.Add(new PnP.Framework.Provisioning.Model.View()
            {
                SchemaXml = @"<View DisplayName=""View Two"">
                  <ViewFields>
                    <FieldRef Name=""ID"" />
                    <FieldRef Name=""Title"" />
                    <FieldRef Name=""ProjectID"" />
                    <FieldRef Name=""ProjectName"" />
                  </ViewFields>
                </View>"
            });
            #endregion

            #region fieldrefs
            list.FieldRefs.Add(new FieldRef("ProjectID")
            {
                Id = new Guid("{23203E97-3BFE-40CB-AFB4-07AA2B86BF45}"),
                DisplayName = "Project ID",
                Hidden = false,
                Required = true
            });
            list.FieldRefs.Add(new FieldRef("ProjectName")
            {
                Id = new Guid("{B01B3DBC-4630-4ED1-B5BA-321BC7841E3D}"),
                DisplayName = "Project Name",
                Hidden = true,
                Required = false
            });
            list.FieldRefs.Add(new FieldRef("ProjectManager")
            {
                Id = new Guid("{A5DE9600-B7A6-42DD-A05E-10D4F1500208}"),
                DisplayName = "Project Manager",
                Hidden = false,
                Required = true
            });
            #endregion

            #region folders
            var folder01 = new PnP.Framework.Provisioning.Model.Folder("Folder01");
            var folder02 = new PnP.Framework.Provisioning.Model.Folder("Folder02");
            folder01.Folders.Add(new PnP.Framework.Provisioning.Model.Folder("Folder01.01",
                security: new PnP.Framework.Provisioning.Model.ObjectSecurity(new List<PnP.Framework.Provisioning.Model.RoleAssignment>() {
                    new PnP.Framework.Provisioning.Model.RoleAssignment()
                    {
                        Principal="Principal01",
                        RoleDefinition ="Read"
                    },
                    new PnP.Framework.Provisioning.Model.RoleAssignment()
                    {
                        Principal="Principal02",
                        RoleDefinition ="Contribute"
                    },
                    new PnP.Framework.Provisioning.Model.RoleAssignment()
                    {
                        Principal="Principal03",
                        RoleDefinition ="FullControl"
                    }
                })
                {
                    CopyRoleAssignments = true,
                    ClearSubscopes = true
                }));
            folder01.Folders.Add(new PnP.Framework.Provisioning.Model.Folder("Folder01.02"));
            var folder03 = new PnP.Framework.Provisioning.Model.Folder("Folder03")
            {
                ContentTypeID = "0x0120001234567890"
            };
            folder03.Properties.Add("CustomProperty01", "CustomValue01");
            folder03.Properties.Add("CustomProperty02", "CustomValue02");

            list.Folders.Add(folder01);
            list.Folders.Add(folder02);
            list.Folders.Add(folder03);
            #endregion

            list.Fields.Add(new PnP.Framework.Provisioning.Model.Field()
            {
                SchemaXml = "<Field ID=\"{23203E97-3BFE-40CB-AFB4-07AA2B86BF45}\" Type=\"Text\" Name=\"ProjectID\" DisplayName=\"Project ID\" Group=\"My Columns\" MaxLength=\"255\" AllowDeletion=\"TRUE\" Required=\"TRUE\" />"
            });
            list.Fields.Add(new PnP.Framework.Provisioning.Model.Field()
            {
                SchemaXml = "<Field ID=\"{B01B3DBC-4630-4ED1-B5BA-321BC7841E3D}\" Type=\"Text\" Name=\"ProjectName\" DisplayName=\"Project Name\" Group=\"My Columns\" MaxLength=\"255\" AllowDeletion=\"TRUE\" />"
            });

            result.Lists.Add(list);

            var serializer = new TargetSerializer();
            provider.SaveAs(result, TEST_OUT_FILE, serializer);

            var path = $"{provider.Connector.Parameters["ConnectionString"]}\\{provider.Connector.Parameters["Container"]}\\{TEST_OUT_FILE}";
            Assert.IsTrue(File.Exists(path)); path = Path.GetFullPath(path);
            var xml = XDocument.Load(path);
            var wrappedResult = XMLSerializer.Deserialize<TargetProvisioning>(xml);

            var template = wrappedResult.Templates[0].ProvisioningTemplate.First();

            Assert.IsNotNull(template.Lists);
            Assert.AreEqual(1, template.Lists.Length);

            var l = template.Lists.FirstOrDefault(ls => ls.Title == "Project Documents");
            Assert.IsNotNull(l);
            Assert.AreEqual("Project Documents are stored here", l.Description);
            Assert.AreEqual("document.dotx", l.DocumentTemplate);
            Assert.IsTrue(l.OnQuickLaunch);
            Assert.AreEqual(101, l.TemplateType);
            Assert.AreEqual("/Lists/ProjectDocuments", l.Url);
            Assert.IsTrue(l.EnableVersioning);
            Assert.IsTrue(l.ForceCheckout);
            Assert.IsTrue(l.RemoveExistingContentTypes);
            Assert.IsTrue(l.NoCrawl);
            Assert.AreEqual(PnP.Framework.Provisioning.Model.ListExperience.ClassicExperience, list.ListExperience);
            Assert.IsTrue(l.ContentTypesEnabled);
            Assert.AreEqual("./Forms/Display.aspx", list.DefaultDisplayFormUrl);
            Assert.AreEqual("./Forms/Edit.aspx", list.DefaultEditFormUrl);
            Assert.AreEqual("./Forms/NewItem.aspx", list.DefaultNewFormUrl);
            Assert.AreEqual(ListReadingDirection.LTR, list.Direction);
            Assert.AreEqual(1, l.DraftVersionVisibility);
            Assert.IsTrue(l.DraftVersionVisibilitySpecified);
            Assert.IsTrue(l.EnableAttachments);
            Assert.IsTrue(l.EnableFolderCreation);
            Assert.IsTrue(l.EnableMinorVersions);
            Assert.IsTrue(l.EnableModeration);
            Assert.IsTrue(l.Hidden);
            Assert.AreEqual("./Images/Sample.png", list.ImageUrl);
            Assert.AreEqual(true, list.IrmExpire);
            Assert.AreEqual(false, list.IrmReject);
            Assert.AreEqual(false, list.IsApplicationList);
            Assert.AreEqual(10, l.MaxVersionLimit);
            Assert.IsTrue(l.MaxVersionLimitSpecified);
            Assert.AreEqual(2, l.MinorVersionLimit);
            Assert.IsTrue(l.MinorVersionLimitSpecified);
            Assert.AreEqual(11, list.ReadSecurity);
            Assert.AreEqual("30FB193E-016E-45A6-B6FD-C6C2B31AA150".ToLower(), l.TemplateFeatureID);
            Assert.AreEqual("sample formula here", list.ValidationFormula);
            Assert.AreEqual("validation message here", list.ValidationMessage);
            Assert.AreEqual("fake-template.stp", list.TemplateInternalName);
            Assert.IsTrue(list.EnableAudienceTargeting);
            Assert.IsTrue(list.EnableClassicAudienceTargeting);

            Assert.AreEqual(120, list.Webhooks[0].ExpiresInDays);
            Assert.AreEqual("http://myapp.azurewebsites.net/WebHookListener", list.Webhooks[0].ServerNotificationUrl);

            Assert.IsNotNull(l.Security);
            var security = l.Security.BreakRoleInheritance;
            Assert.IsNotNull(security);
            Assert.IsTrue(security.ClearSubscopes);
            Assert.IsTrue(security.CopyRoleAssignments);
            Assert.IsNotNull(security.RoleAssignment);
            Assert.AreEqual(3, security.RoleAssignment.Length);
            var ra = security.RoleAssignment.FirstOrDefault(r => r.Principal == "Principal01");
            Assert.IsNotNull(ra);
            Assert.AreEqual("Read", ra.RoleDefinition);
            ra = security.RoleAssignment.FirstOrDefault(r => r.Principal == "Principal02");
            Assert.IsNotNull(ra);
            Assert.AreEqual("Contribute", ra.RoleDefinition);
            ra = security.RoleAssignment.FirstOrDefault(r => r.Principal == "Principal03");
            Assert.IsNotNull(ra);
            Assert.AreEqual("FullControl", ra.RoleDefinition);

            Assert.IsNotNull(list.ContentTypeBindings);
            Assert.AreEqual(3, list.ContentTypeBindings.Count);
            var ct = l.ContentTypeBindings.FirstOrDefault(c => c.ContentTypeID == "0x01005D4F34E4BE7F4B6892AEBE088EDD215E");
            Assert.IsNotNull(ct);
            Assert.IsTrue(ct.Default);
            Assert.IsFalse(ct.Remove);
            ct = l.ContentTypeBindings.FirstOrDefault(c => c.ContentTypeID == "0x0101");
            Assert.IsNotNull(ct);
            Assert.IsFalse(ct.Default);
            Assert.IsTrue(ct.Remove);
            ct = l.ContentTypeBindings.FirstOrDefault(c => c.ContentTypeID == "0x0102");
            Assert.IsNotNull(ct);
            Assert.IsFalse(ct.Default);
            Assert.IsFalse(ct.Remove);

            Assert.IsNotNull(list.DefaultColumnValues);
            Assert.AreEqual(2, list.DefaultColumnValues.Count);
            Assert.AreEqual("Custom Value 1", list.DefaultColumnValues["Field1"]);
            Assert.AreEqual("Custom Value 2", list.DefaultColumnValues["Field2"]);

            Assert.IsNotNull(l.FieldDefaults);
            Assert.AreEqual(4, l.FieldDefaults.Length);
            var fd = l.FieldDefaults.FirstOrDefault(f => f.FieldName == "Field01");
            Assert.IsNotNull(fd);
            Assert.AreEqual("DefaultValue01", fd.Value);
            fd = l.FieldDefaults.FirstOrDefault(f => f.FieldName == "Field02");
            Assert.IsNotNull(fd);
            Assert.AreEqual("DefaultValue02", fd.Value);
            fd = l.FieldDefaults.FirstOrDefault(f => f.FieldName == "Field03");
            Assert.IsNotNull(fd);
            Assert.AreEqual("DefaultValue03", fd.Value);
            fd = l.FieldDefaults.FirstOrDefault(f => f.FieldName == "Field04");
            Assert.IsNotNull(fd);
            Assert.AreEqual("DefaultValue04", fd.Value);

            Assert.IsNotNull(l.DataRows);
            Assert.AreEqual(3, l.DataRows.DataRow.Length);
            #region data row 1 asserts
            var dr = l.DataRows.DataRow.FirstOrDefault(r => r.DataValue.Any(d => d.Value.StartsWith("Value01")));
            Assert.IsNotNull(dr);
            Assert.IsNotNull(dr.Security);
            security = dr.Security.BreakRoleInheritance;
            Assert.IsNotNull(security);
            Assert.IsTrue(security.ClearSubscopes);
            Assert.IsTrue(security.CopyRoleAssignments);
            Assert.IsNotNull(security.RoleAssignment);
            Assert.AreEqual(3, security.RoleAssignment.Length);
            ra = security.RoleAssignment.FirstOrDefault(r => r.Principal == "Principal01");
            Assert.IsNotNull(ra);
            Assert.AreEqual("Read", ra.RoleDefinition);
            ra = security.RoleAssignment.FirstOrDefault(r => r.Principal == "Principal02");
            Assert.IsNotNull(ra);
            Assert.AreEqual("Contribute", ra.RoleDefinition);
            ra = security.RoleAssignment.FirstOrDefault(r => r.Principal == "Principal03");
            Assert.IsNotNull(ra);
            Assert.AreEqual("FullControl", ra.RoleDefinition);

            Assert.IsNotNull(dr.DataValue);
            var dv = dr.DataValue.FirstOrDefault(d => d.FieldName == "Field01");
            Assert.IsNotNull(dv);
            Assert.AreEqual("Value01-01", dv.Value);
            dv = dr.DataValue.FirstOrDefault(d => d.FieldName == "Field02");
            Assert.IsNotNull(dv);
            Assert.AreEqual("Value01-02", dv.Value);
            dv = dr.DataValue.FirstOrDefault(d => d.FieldName == "Field03");
            Assert.IsNotNull(dv);
            Assert.AreEqual("Value01-03", dv.Value);
            dv = dr.DataValue.FirstOrDefault(d => d.FieldName == "Field04");
            Assert.IsNotNull(dv);
            Assert.AreEqual("Value01-04", dv.Value);
            #endregion
            #region data row 2 asserts
            dr = l.DataRows.DataRow.FirstOrDefault(r => r.DataValue.Any(d => d.Value.StartsWith("Value02")));
            Assert.IsNotNull(dr);
            Assert.IsNotNull(dr.Security);
            security = dr.Security.BreakRoleInheritance;
            Assert.IsNotNull(security);
            Assert.IsFalse(security.ClearSubscopes);
            Assert.IsFalse(security.CopyRoleAssignments);
            Assert.IsNotNull(security.RoleAssignment);
            Assert.AreEqual(3, security.RoleAssignment.Length);
            ra = security.RoleAssignment.FirstOrDefault(r => r.Principal == "Principal01");
            Assert.IsNotNull(ra);
            Assert.AreEqual("Read", ra.RoleDefinition);
            ra = security.RoleAssignment.FirstOrDefault(r => r.Principal == "Principal02");
            Assert.IsNotNull(ra);
            Assert.AreEqual("Contribute", ra.RoleDefinition);
            ra = security.RoleAssignment.FirstOrDefault(r => r.Principal == "Principal03");
            Assert.IsNotNull(ra);
            Assert.AreEqual("FullControl", ra.RoleDefinition);

            Assert.IsNotNull(dr.DataValue);
            dv = dr.DataValue.FirstOrDefault(d => d.FieldName == "Field01");
            Assert.IsNotNull(dv);
            Assert.AreEqual("Value02-01", dv.Value);
            dv = dr.DataValue.FirstOrDefault(d => d.FieldName == "Field02");
            Assert.IsNotNull(dv);
            Assert.AreEqual("Value02-02", dv.Value);
            dv = dr.DataValue.FirstOrDefault(d => d.FieldName == "Field03");
            Assert.IsNotNull(dv);
            Assert.AreEqual("Value02-03", dv.Value);
            dv = dr.DataValue.FirstOrDefault(d => d.FieldName == "Field04");
            Assert.IsNotNull(dv);
            Assert.AreEqual("Value02-04", dv.Value);
            #endregion
            #region data row 3 asserts
            dr = l.DataRows.DataRow.FirstOrDefault(r => r.DataValue.Any(d => d.Value.StartsWith("Value03")));
            Assert.IsNotNull(dr);
            Assert.IsNull(dr.Security);

            Assert.IsNotNull(dr.DataValue);
            dv = dr.DataValue.FirstOrDefault(d => d.FieldName == "Field01");
            Assert.IsNotNull(dv);
            Assert.AreEqual("Value03-01", dv.Value);
            dv = dr.DataValue.FirstOrDefault(d => d.FieldName == "Field02");
            Assert.IsNotNull(dv);
            Assert.AreEqual("Value03-02", dv.Value);
            dv = dr.DataValue.FirstOrDefault(d => d.FieldName == "Field03");
            Assert.IsNotNull(dv);
            Assert.AreEqual("Value03-03", dv.Value);
            dv = dr.DataValue.FirstOrDefault(d => d.FieldName == "Field04");
            Assert.IsNotNull(dv);
            Assert.AreEqual("Value03-04", dv.Value);
            #endregion

            #region user custom action
            Assert.IsNotNull(l.UserCustomActions);
            Assert.AreEqual(1, l.UserCustomActions.Length);
            var ua = l.UserCustomActions.FirstOrDefault(a => a.Name == "SampleCustomAction");
            Assert.IsNotNull(ua);
            Assert.AreEqual("Just a sample custom action", ua.Description);
            Assert.IsTrue(ua.Enabled);
            Assert.AreEqual("Samples", ua.Group);
            Assert.AreEqual("OneImage.png", ua.ImageUrl);
            Assert.AreEqual("Any", ua.Location);
            Assert.AreEqual("0x0101", ua.RegistrationId);
            Assert.AreEqual(RegistrationType.ContentType, ua.RegistrationType);
            Assert.AreEqual(100, ua.Sequence);
            Assert.AreEqual("scriptblock", ua.ScriptBlock);
            Assert.AreEqual("script.js", ua.ScriptSrc);
            Assert.AreEqual("http://somewhere.com/", ua.Url);
            Assert.AreEqual("Sample Action", ua.Title);
            Assert.IsTrue(ua.Remove);
            Assert.IsNotNull(ua.CommandUIExtension);
            Assert.IsNotNull(ua.CommandUIExtension.Any);
            Assert.AreEqual(1, ua.CommandUIExtension.Any.Length);
            Assert.IsNotNull(ua.Rights);
            Assert.IsTrue(ua.Rights.Contains("AddListItems"));
            #endregion

            Assert.IsNotNull(l.Views);
            Assert.IsNotNull(l.Views.Any);
            Assert.AreEqual(2, l.Views.Any.Length);

            #region field refs
            Assert.IsNotNull(l.FieldRefs);
            Assert.AreEqual(3, l.FieldRefs.Length);
            var fr = l.FieldRefs.FirstOrDefault(f => f.Name == "ProjectID");
            Assert.IsNotNull(fr);
            Assert.AreEqual("23203E97-3BFE-40CB-AFB4-07AA2B86BF45".ToLower(), fr.ID);
            Assert.AreEqual("Project ID", fr.DisplayName);
            Assert.IsFalse(fr.Hidden);
            Assert.IsTrue(fr.Required);
            fr = l.FieldRefs.FirstOrDefault(f => f.Name == "ProjectName");
            Assert.IsNotNull(fr);
            Assert.AreEqual("B01B3DBC-4630-4ED1-B5BA-321BC7841E3D".ToLower(), fr.ID);
            Assert.AreEqual("Project Name", fr.DisplayName);
            Assert.IsTrue(fr.Hidden);
            Assert.IsFalse(fr.Required);
            fr = l.FieldRefs.FirstOrDefault(f => f.Name == "ProjectManager");
            Assert.IsNotNull(fr);
            Assert.AreEqual("A5DE9600-B7A6-42DD-A05E-10D4F1500208".ToLower(), fr.ID);
            Assert.AreEqual("Project Manager", fr.DisplayName);
            Assert.IsFalse(fr.Hidden);
            Assert.IsTrue(fr.Required);
            #endregion

            #region folders
            Assert.IsNotNull(l.Folders);
            Assert.AreEqual(3, l.Folders.Length);
            var fl = l.Folders.FirstOrDefault(f => f.Name == "Folder02");
            Assert.IsNotNull(fl);
            Assert.IsNull(fl.Folder1);
            fl = l.Folders.FirstOrDefault(f => f.Name == "Folder01");
            Assert.IsNotNull(fl);
            Assert.IsNotNull(fl.Folder1);
            var fl1 = fl.Folder1.FirstOrDefault(f => f.Name == "Folder01.02");
            Assert.IsNotNull(fl1);
            Assert.IsNull(fl1.Folder1);
            fl1 = fl.Folder1.FirstOrDefault(f => f.Name == "Folder01.01");
            Assert.IsNull(fl1.Folder1);
            Assert.IsNotNull(fl1.Security);
            security = fl1.Security.BreakRoleInheritance;
            Assert.IsNotNull(security);
            Assert.IsTrue(security.ClearSubscopes);
            Assert.IsTrue(security.CopyRoleAssignments);
            Assert.IsNotNull(security.RoleAssignment);
            Assert.AreEqual(3, security.RoleAssignment.Length);
            ra = security.RoleAssignment.FirstOrDefault(r => r.Principal == "Principal01");
            Assert.IsNotNull(ra);
            Assert.AreEqual("Read", ra.RoleDefinition);
            ra = security.RoleAssignment.FirstOrDefault(r => r.Principal == "Principal02");
            Assert.IsNotNull(ra);
            Assert.AreEqual("Contribute", ra.RoleDefinition);
            ra = security.RoleAssignment.FirstOrDefault(r => r.Principal == "Principal03");
            Assert.IsNotNull(ra);
            Assert.AreEqual("FullControl", ra.RoleDefinition);
            fl = l.Folders.FirstOrDefault(f => f.Name == "Folder03");
            Assert.IsNotNull(fl);
            Assert.AreEqual("0x0120001234567890", fl.ContentTypeID);
            Assert.IsNotNull(fl.Properties);
            Assert.AreEqual(2, fl.Properties.Length);
            Assert.AreEqual("CustomProperty01", fl.Properties[0].Key);
            Assert.AreEqual("CustomValue01", fl.Properties[0].Value);
            Assert.AreEqual("CustomProperty02", fl.Properties[1].Key);
            Assert.AreEqual("CustomValue02", fl.Properties[1].Value);

            #endregion

            #region IRM Settings
            Assert.AreEqual(true, list.IRMSettings.AllowPrint);
            Assert.AreEqual(false, list.IRMSettings.AllowWriteCopy);
            Assert.AreEqual(true, list.IRMSettings.AllowScript);
            Assert.AreEqual(true, list.IRMSettings.DisableDocumentBrowserView);
            Assert.AreEqual(30, list.IRMSettings.DocumentAccessExpireDays);
            Assert.AreEqual(90, list.IRMSettings.DocumentLibraryProtectionExpiresInDays);
            Assert.AreEqual(true, list.IRMSettings.Enabled);
            Assert.AreEqual(true, list.IRMSettings.EnableDocumentAccessExpire);
            Assert.AreEqual(true, list.IRMSettings.EnableDocumentBrowserPublishingView);
            Assert.AreEqual(false, list.IRMSettings.EnableGroupProtection);
            Assert.AreEqual(false, list.IRMSettings.EnableLicenseCacheExpire);
            Assert.AreEqual("Custom IRM Group", list.IRMSettings.GroupName);
            Assert.AreEqual(60, list.IRMSettings.LicenseCacheExpireDays);
            Assert.AreEqual("Sample IRM Policy", list.IRMSettings.PolicyDescription);
            Assert.AreEqual("Sample IRM Policy", list.IRMSettings.PolicyTitle);
            #endregion

            Assert.IsNotNull(l.Fields);
            Assert.IsNotNull(l.Fields.Any);
            Assert.AreEqual(2, l.Fields.Any.Length);
            Assert.IsTrue(l.Fields.Any.All(x => x.OuterXml.StartsWith("<Field")));
        }

        [TestMethod]
        [TestCategory(TEST_CATEGORY)]
        public void XMLSerializer_Deserialize_Features()
        {
            var provider = new XMLFileSystemTemplateProvider($@"{AppDomain.CurrentDomain.BaseDirectory}\..\..\..\Resources", "Templates");

            var serializer = new TargetSerializer();
            var template = provider.GetTemplate(TEST_TEMPLATE, serializer);
            Assert.IsNotNull(template.Features);
            Assert.IsNotNull(template.Features.SiteFeatures);
            Assert.AreEqual(3, template.Features.SiteFeatures.Count);
            var feature = template.Features.SiteFeatures.FirstOrDefault(f => f.Id == new Guid("b50e3104-6812-424f-a011-cc90e6327318"));
            Assert.IsNotNull(feature);
            Assert.IsFalse(feature.Deactivate);
            feature = template.Features.SiteFeatures.FirstOrDefault(f => f.Id == new Guid("9c0834e1-ba47-4d49-812b-7d4fb6fea211"));
            Assert.IsNotNull(feature);
            Assert.IsFalse(feature.Deactivate);
            feature = template.Features.SiteFeatures.FirstOrDefault(f => f.Id == new Guid("0af5989a-3aea-4519-8ab0-85d91abe39ff"));
            Assert.IsNotNull(feature);
            Assert.IsTrue(feature.Deactivate);

            Assert.IsNotNull(template.Features.WebFeatures);
            Assert.AreEqual(4, template.Features.WebFeatures.Count);
            feature = template.Features.WebFeatures.FirstOrDefault(f => f.Id == new Guid("7201d6a4-a5d3-49a1-8c19-19c4bac6e668"));
            Assert.IsNotNull(feature);
            Assert.IsFalse(feature.Deactivate);
            feature = template.Features.WebFeatures.FirstOrDefault(f => f.Id == new Guid("961d6a9c-4388-4cf2-9733-38ee8c89afd4"));
            Assert.IsNotNull(feature);
            Assert.IsFalse(feature.Deactivate);
            feature = template.Features.WebFeatures.FirstOrDefault(f => f.Id == new Guid("e2f2bb18-891d-4812-97df-c265afdba297"));
            Assert.IsNotNull(feature);
            Assert.IsFalse(feature.Deactivate);
            feature = template.Features.WebFeatures.FirstOrDefault(f => f.Id == new Guid("4aec7207-0d02-4f4f-aa07-b370199cd0c7"));
            Assert.IsNotNull(feature);
            Assert.IsTrue(feature.Deactivate);
        }

        [TestMethod]
        [TestCategory(TEST_CATEGORY)]
        public void XMLSerializer_Serialize_Features()
        {
            var provider = new XMLFileSystemTemplateProvider($@"{AppDomain.CurrentDomain.BaseDirectory}\..\..\..\Resources", "Templates");

            var result = new ProvisioningTemplate
            {
                Features = new PnP.Framework.Provisioning.Model.Features()
            };

            result.Features.SiteFeatures.Add(new PnP.Framework.Provisioning.Model.Feature()
            {
                Id = new Guid("d8f187e3-2bf3-43a3-99a0-024edaffab5e")
            });
            result.Features.SiteFeatures.Add(new PnP.Framework.Provisioning.Model.Feature()
            {
                Id = new Guid("89c029c5-d289-4936-8ba6-6f3386a8a03f"),
                Deactivate = true
            });
            result.Features.WebFeatures.Add(new PnP.Framework.Provisioning.Model.Feature()
            {
                Id = new Guid("a22d7848-6d17-47b5-9c1c-cecc98a6b258")
            });
            result.Features.WebFeatures.Add(new PnP.Framework.Provisioning.Model.Feature()
            {
                Id = new Guid("d60aed53-05f3-4d1c-a12f-677da19a8c31"),
                Deactivate = true
            });

            var serializer = new TargetSerializer();
            provider.SaveAs(result, TEST_OUT_FILE, serializer);

            var path = $"{provider.Connector.Parameters["ConnectionString"]}\\{provider.Connector.Parameters["Container"]}\\{TEST_OUT_FILE}";
            Assert.IsTrue(File.Exists(path)); path = Path.GetFullPath(path);
            var xml = XDocument.Load(path);
            var wrappedResult =
                XMLSerializer.Deserialize<TargetProvisioning>(xml);

            var template = wrappedResult.Templates[0].ProvisioningTemplate.First();

            Assert.IsNotNull(template.Features);
            Assert.IsNotNull(template.Features.SiteFeatures);
            Assert.AreEqual(2, template.Features.SiteFeatures.Length);
            var feature = template.Features.SiteFeatures.FirstOrDefault(f => f.ID == "d8f187e3-2bf3-43a3-99a0-024edaffab5e");
            Assert.IsNotNull(feature);
            Assert.IsFalse(feature.Deactivate);
            feature = template.Features.SiteFeatures.FirstOrDefault(f => f.ID == "89c029c5-d289-4936-8ba6-6f3386a8a03f");
            Assert.IsNotNull(feature);
            Assert.IsTrue(feature.Deactivate);

            Assert.IsNotNull(template.Features.WebFeatures);
            Assert.AreEqual(2, template.Features.WebFeatures.Length);
            feature = template.Features.WebFeatures.FirstOrDefault(f => f.ID == "a22d7848-6d17-47b5-9c1c-cecc98a6b258");
            Assert.IsNotNull(feature);
            Assert.IsFalse(feature.Deactivate);
            feature = template.Features.WebFeatures.FirstOrDefault(f => f.ID == "d60aed53-05f3-4d1c-a12f-677da19a8c31");
            Assert.IsNotNull(feature);
            Assert.IsTrue(feature.Deactivate);
        }

        [TestMethod]
        [TestCategory(TEST_CATEGORY)]
        public void XMLSerializer_Deserialize_CustomActions()
        {
            var provider = new XMLFileSystemTemplateProvider($@"{AppDomain.CurrentDomain.BaseDirectory}\..\..\..\Resources", "Templates");

            var serializer = new TargetSerializer();
            var template = provider.GetTemplate(TEST_TEMPLATE, serializer);

            Assert.IsNotNull(template);
            Assert.IsNotNull(template.CustomActions);
            Assert.IsNotNull(template.CustomActions.SiteCustomActions);
            Assert.IsNotNull(template.CustomActions.WebCustomActions);

            var ca = template.CustomActions.SiteCustomActions.FirstOrDefault(c => c.Name == "CA_SITE_SETTINGS_SITECLASSIFICATION");
            Assert.IsNotNull(ca);
            Assert.AreEqual("Site Classification Application", ca.Description);
            Assert.AreEqual("Microsoft.SharePoint.SiteSettings", ca.Location);
            Assert.AreEqual("Site Classification", ca.Title);
            Assert.AreEqual(1000, ca.Sequence);
            Assert.IsTrue(ca.Rights.Has(PermissionKind.ManageWeb));
            Assert.AreEqual("https://spmanaged.azurewebsites.net/pages/index.aspx?SPHostUrl={0}", ca.Url);
            Assert.AreEqual(UserCustomActionRegistrationType.None, ca.RegistrationType);
            Assert.AreEqual(new Guid("70056b07-3a88-4882-989f-eb2f1a74ff5f"), ca.ClientSideComponentId);
            Assert.AreEqual("{}", ca.ClientSideComponentProperties);

            ca = template.CustomActions.WebCustomActions.FirstOrDefault(c => c.Name == "CA_SUBSITE_OVERRIDE");
            Assert.IsNotNull(ca);
            Assert.AreEqual("Override new sub-site link", ca.Description);
            Assert.AreEqual("Microsoft.SharePoint.SiteSettings", ca.Location);
            Assert.AreEqual("SubSite Overide", ca.Title);
            Assert.AreEqual(1000, ca.Sequence);
        }

        [TestMethod]
        [TestCategory(TEST_CATEGORY)]
        public void XMLSerializer_Serialize_CustomActions()
        {
            var provider = new XMLFileSystemTemplateProvider($@"{AppDomain.CurrentDomain.BaseDirectory}\..\..\..\Resources", "Templates");

            var result = new ProvisioningTemplate
            {
                CustomActions = new PnP.Framework.Provisioning.Model.CustomActions()
            };

            var can = new PnP.Framework.Provisioning.Model.CustomAction()
            {
                Name = "CA_SITE_SETTINGS_SITECLASSIFICATION",
                Description = "Site Classification Application",
                Location = "Microsoft.SharePoint.SiteSettings",
                Title = "Site Classification",
                Url = "https://spmanaged.azurewebsites.net/pages/index.aspx?SPHostUrl={0}",
                Sequence = 1000,
                RegistrationType = UserCustomActionRegistrationType.None,
                Rights = new BasePermissions(),
                ImageUrl = "http://sharepoint.com",
                RegistrationId = "101",
                ScriptBlock = "alert('boo')",
                CommandUIExtension = XElement.Parse(@"<CommandUIExtension><CommandUIDefinitions>
                <CommandUIDefinition Location=""Ribbon.Documents.Copies.Controls._children"">
                  <Button Sequence = ""15"" TemplateAlias = ""o1"" ToolTipDescription = ""Download all files separately"" ToolTipTitle = ""Download All"" Description = ""Download all files separately"" LabelText = ""Download All"" Image32by32 = ""~sitecollection/SiteAssets/DownloadAll32x32.png"" Image16by16 = ""~sitecollection/SiteAssets/DownloadAll16x16.png"" Command = ""OfficeDevPnP.Cmd.DownloadAll"" Id = ""Ribbon.Documents.Copies.OfficeDevPnPDownloadAll"" />
                </CommandUIDefinition>
                <CommandUIDefinition Location = ""Ribbon.Documents.Copies.Controls._children"">
                  <Button Sequence = ""20"" TemplateAlias = ""o1"" ToolTipDescription = ""Download all files as single Zip archive"" ToolTipTitle = ""Download All as Zip"" Description = ""Download all files as single Zip"" LabelText = ""Download All as Zip"" Image32by32 = ""~sitecollection/SiteAssets/DownloadAllAsZip32x32.png"" Image16by16 = ""~sitecollection/SiteAssets/DownloadAllAsZip16x16.png"" Command = ""OfficeDevPnP.Cmd.DownloadAllAsZip"" Id = ""Ribbon.Documents.Copies.OfficeDevPnPDownloadAllAsZip"" />
                </CommandUIDefinition>
              </CommandUIDefinitions>
              <CommandUIHandlers>
                <CommandUIHandler Command = ""OfficeDevPnP.Cmd.DownloadAll"" EnabledScript = ""javascript:PnP.Framework.RibbonManager.isListViewButtonEnabled('DownloadAll');"" CommandAction = ""javascript:PnP.Framework.RibbonManager.invokeCommand('DownloadAll');"" />
                <CommandUIHandler Command = ""OfficeDevPnP.Cmd.DownloadAllAsZip"" EnabledScript = ""javascript:PnP.Framework.RibbonManager.isListViewButtonEnabled('DownloadAllAsZip');"" CommandAction = ""javascript:PnP.Framework.RibbonManager.invokeCommand('DownloadAllAsZip');"" />
              </CommandUIHandlers></CommandUIExtension>"),
            };
            can.Rights.Set(PermissionKind.ManageWeb);
            result.CustomActions.SiteCustomActions.Add(can);

            can = new PnP.Framework.Provisioning.Model.CustomAction()
            {
                Name = "CA_SUBSITE_OVERRIDE",
                Description = "Override new sub-site link",
                Location = "ScriptLink",
                Title = "SubSiteOveride",
                Sequence = 100,
                ScriptSrc = "~site/PnP_Provisioning_JS/PnP_EmbeddedJS.js",
                RegistrationType = UserCustomActionRegistrationType.ContentType,
                ClientSideComponentId = new Guid("70056b07-3a88-4882-989f-eb2f1a74ff5f"),
                ClientSideComponentProperties = "{}",
            };
            result.CustomActions.SiteCustomActions.Add(can);

            can = new PnP.Framework.Provisioning.Model.CustomAction()
            {
                Name = "CA_WEB_DOCLIB_MENU_SAMPLE",
                Description = "Document Library Custom Menu",
                Group = "ActionsMenu",
                Location = "Microsoft.SharePoint.StandardMenu",
                Title = "DocLib Custom Menu",
                Sequence = 100,
                Url = "/_layouts/CustomActionsHello.aspx?ActionsMenu",
                RegistrationType = UserCustomActionRegistrationType.None
            };
            result.CustomActions.WebCustomActions.Add(can);

            var serializer = new TargetSerializer();
            provider.SaveAs(result, TEST_OUT_FILE, serializer);

            var path = $"{provider.Connector.Parameters["ConnectionString"]}\\{provider.Connector.Parameters["Container"]}\\{TEST_OUT_FILE}";
            Assert.IsTrue(File.Exists(path)); path = Path.GetFullPath(path);
            var xml = XDocument.Load(path);
            var wrappedResult =
                XMLSerializer.Deserialize<TargetProvisioning>(xml);

            var template = wrappedResult.Templates[0].ProvisioningTemplate.First();

            Assert.IsNotNull(template.CustomActions.SiteCustomActions);
            Assert.IsNotNull(template.CustomActions.WebCustomActions);

            var ca = template.CustomActions.SiteCustomActions.FirstOrDefault(c => c.Name == "CA_SITE_SETTINGS_SITECLASSIFICATION");
            Assert.IsNotNull(ca);
            Assert.AreEqual("Site Classification Application", ca.Description);
            Assert.AreEqual("Microsoft.SharePoint.SiteSettings", ca.Location);
            Assert.AreEqual("Site Classification", ca.Title);
            Assert.AreEqual(1000, ca.Sequence);
            Assert.AreEqual("ManageWeb", ca.Rights);
            Assert.AreEqual("https://spmanaged.azurewebsites.net/pages/index.aspx?SPHostUrl={0}", ca.Url);
            Assert.AreEqual(RegistrationType.None, ca.RegistrationType);
            Assert.IsNotNull(ca.CommandUIExtension);
            Assert.AreEqual("http://sharepoint.com", ca.ImageUrl);
            Assert.AreEqual("101", ca.RegistrationId);
            Assert.AreEqual("alert('boo')", ca.ScriptBlock);
            Assert.IsNotNull(ca.CommandUIExtension);
            Assert.IsNotNull(ca.CommandUIExtension.Any);
            Assert.AreEqual(2, ca.CommandUIExtension.Any.Length);

            ca = template.CustomActions.SiteCustomActions.FirstOrDefault(c => c.Name == "CA_SUBSITE_OVERRIDE");
            Assert.IsNotNull(ca);
            Assert.AreEqual("Override new sub-site link", ca.Description);
            Assert.AreEqual("ScriptLink", ca.Location);
            Assert.AreEqual("SubSiteOveride", ca.Title);
            Assert.AreEqual(100, ca.Sequence);
            Assert.AreEqual("~site/PnP_Provisioning_JS/PnP_EmbeddedJS.js", ca.ScriptSrc);
            Assert.AreEqual(RegistrationType.ContentType, ca.RegistrationType);
            Assert.AreEqual("70056b07-3a88-4882-989f-eb2f1a74ff5f", ca.ClientSideComponentId);
            Assert.AreEqual("{}", ca.ClientSideComponentProperties);
            Assert.IsNull(ca.CommandUIExtension);

            ca = template.CustomActions.WebCustomActions.FirstOrDefault(c => c.Name == "CA_WEB_DOCLIB_MENU_SAMPLE");
            Assert.IsNotNull(ca);
            Assert.AreEqual("Document Library Custom Menu", ca.Description);
            Assert.AreEqual("ActionsMenu", ca.Group);
            Assert.AreEqual("Microsoft.SharePoint.StandardMenu", ca.Location);
            Assert.AreEqual("DocLib Custom Menu", ca.Title);
            Assert.AreEqual(100, ca.Sequence);
            Assert.AreEqual("/_layouts/CustomActionsHello.aspx?ActionsMenu", ca.Url);
            Assert.AreEqual(RegistrationType.None, ca.RegistrationType);
            Assert.IsNull(ca.CommandUIExtension);
        }

        [TestMethod]
        [TestCategory(TEST_CATEGORY)]
        public void XMLSerializer_Deserialize_Files()
        {
            var provider = new XMLFileSystemTemplateProvider($@"{AppDomain.CurrentDomain.BaseDirectory}\..\..\..\Resources", "Templates");

            var serializer = new TargetSerializer();
            var template = provider.GetTemplate(TEST_TEMPLATE, serializer);

            Assert.IsNotNull(template);
            Assert.IsNotNull(template.Files);

            var file = template.Files.FirstOrDefault(f => f.Src == "Logo.png");
            Assert.AreEqual("SiteAssets", file.Folder);
            Assert.AreEqual(true, file.Overwrite);
            Assert.AreEqual("CompanyLogo.png", file.TargetFileName);

            file = template.Files.SingleOrDefault(f => f.Src == "CustomPage.aspx");
            Assert.AreEqual(true, file.Security.CopyRoleAssignments);
            Assert.AreEqual(true, file.Security.ClearSubscopes);
            Assert.AreEqual("Power Users", file.Security.RoleAssignments[0].Principal);
            Assert.AreEqual("Full Control", file.Security.RoleAssignments[0].RoleDefinition);

            file = template.Files.SingleOrDefault(f => f.Src == "CustomMaster.master");
            Assert.AreEqual(PnP.Framework.Provisioning.Model.FileLevel.Published, file.Level);

            var dir = template.Directories.SingleOrDefault(d => d.Src == @"c:\LocalPath\StyleLibrary");
            Assert.AreEqual("Style%20Library", dir.Folder);
            Assert.AreEqual(true, dir.Overwrite);
            Assert.AreEqual(true, dir.Recursive);
            Assert.AreEqual("*.css,*.js,*.woff", dir.IncludedExtensions);
            Assert.AreEqual("*.xml,*.json", dir.ExcludedExtensions);
            Assert.AreEqual(@".\InvoicesMetadata.json", template.Directories[1].MetadataMappingFile);
        }

        [TestMethod]
        [TestCategory(TEST_CATEGORY)]
        public void XMLSerializer_Serialize_Files()
        {
            var provider = new XMLFileSystemTemplateProvider($@"{AppDomain.CurrentDomain.BaseDirectory}\..\..\..\Resources", "Templates");

            var result = new ProvisioningTemplate();

            result.Files.Add(new PnP.Framework.Provisioning.Model.File
            {
                Src = "Logo.png",
                Overwrite = true,
                Folder = "SiteAssets",
                TargetFileName = "CompanyLogo.png"
            });
            result.Files.Add(new PnP.Framework.Provisioning.Model.File
            {
                Src = "CustomPage.aspx",
                Overwrite = true,
                Folder = "SitePages",
                Security = new PnP.Framework.Provisioning.Model.ObjectSecurity
                {
                    ClearSubscopes = true,
                    CopyRoleAssignments = true,
                    RoleAssignments =
                    {
                        new PnP.Framework.Provisioning.Model.RoleAssignment
                        {
                            Principal = "Power Users",
                            RoleDefinition = "Full Control"
                        }
                    }
                }
            });
            result.Files.Add(new PnP.Framework.Provisioning.Model.File
            {
                Src = "CustomMaster.master",
                Overwrite = true,
                Folder = "_catalogs/MasterPage",
                Level = FileLevel.Published
            });

            result.Directories.Add(new PnP.Framework.Provisioning.Model.Directory
            {
                Src = @"c:\LocalPath\StyleLibrary",
                Overwrite = true,
                IncludedExtensions = "*.css,*.js,*.woff",
                ExcludedExtensions = "*.xml,*.json",
                Recursive = true,
                Folder = "Style%20Library"
            });

            result.Directories.Add(new PnP.Framework.Provisioning.Model.Directory
            {
                Src = @".\Invoices",
                Overwrite = true,
                Recursive = true,
                Folder = "Invoices",
                MetadataMappingFile = @".\InvoicesMetadata.json"
            });

            result.Directories.Add(new PnP.Framework.Provisioning.Model.Directory
            {
                Src = @"c:\LocalPath\Pages",
                Overwrite = true,
                IncludedExtensions = "*.css,*.js,*.woff",
                ExcludedExtensions = "*.xml,*.json",
                Recursive = true,
                Folder = "{PagesLibrary}",
                Level = FileLevel.Published,
                MetadataMappingFile = @"c:\LocalPath\PagesMetadata.json"
            });

            var serializer = new TargetSerializer();
            provider.SaveAs(result, TEST_OUT_FILE, serializer);

            var path = $"{provider.Connector.Parameters["ConnectionString"]}\\{provider.Connector.Parameters["Container"]}\\{TEST_OUT_FILE}";
            Assert.IsTrue(File.Exists(path)); path = Path.GetFullPath(path);
            var xml = XDocument.Load(path);
            var wrappedResult =
                XMLSerializer.Deserialize<TargetProvisioning>(xml);

            var template = wrappedResult.Templates[0].ProvisioningTemplate.First();

            Assert.IsNotNull(template);
            Assert.IsNotNull(template.Files);

            var file = template.Files.File.FirstOrDefault(f => f.Src == "Logo.png");
            Assert.AreEqual("SiteAssets", file.Folder);
            Assert.AreEqual(true, file.Overwrite);
            Assert.AreEqual("CompanyLogo.png", file.TargetFileName);

            file = template.Files.File.SingleOrDefault(f => f.Src == "CustomPage.aspx");
            Assert.AreEqual(true, file.Security.BreakRoleInheritance.CopyRoleAssignments);
            Assert.AreEqual(true, file.Security.BreakRoleInheritance.ClearSubscopes);
            Assert.AreEqual("Power Users", file.Security.BreakRoleInheritance.RoleAssignment[0].Principal);
            Assert.AreEqual("Full Control", file.Security.BreakRoleInheritance.RoleAssignment[0].RoleDefinition);

            file = template.Files.File.SingleOrDefault(f => f.Src == "CustomMaster.master");
            Assert.AreEqual(TargetNamespace.FileLevel.Published, file.Level);

            var dir = template.Files.Directory.SingleOrDefault(d => d.Src == @"c:\LocalPath\StyleLibrary");
            Assert.AreEqual("Style%20Library", dir.Folder);
            Assert.AreEqual(true, dir.Overwrite);
            Assert.AreEqual(true, dir.Recursive);
            Assert.AreEqual("*.css,*.js,*.woff", dir.IncludedExtensions);
            Assert.AreEqual("*.xml,*.json", dir.ExcludedExtensions);
            Assert.AreEqual(@".\InvoicesMetadata.json", template.Files.Directory[1].MetadataMappingFile);
        }

        [TestMethod]
        [TestCategory(TEST_CATEGORY)]
        public void XMLSerializer_Deserialize_Pages()
        {
            var provider = new XMLFileSystemTemplateProvider($@"{AppDomain.CurrentDomain.BaseDirectory}\..\..\..\Resources", "Templates");

            var serializer = new TargetSerializer();
            var template = provider.GetTemplate(TEST_TEMPLATE, serializer);
            var pages = template.Pages;

            Assert.AreEqual(WikiPageLayout.TwoColumns, pages[0].Layout);
            Assert.AreEqual(true, template.Pages[0].Overwrite);
            Assert.AreEqual("{site}/SitePages/DemoWikiPage.aspx", pages[0].Url);

            Assert.AreEqual(true, pages[1].Security.CopyRoleAssignments);
            Assert.AreEqual(true, pages[1].Security.ClearSubscopes);
            Assert.AreEqual("Power Users", pages[1].Security.RoleAssignments[0].Principal);
            Assert.AreEqual("Full Control", pages[1].Security.RoleAssignments[0].RoleDefinition);
            Assert.AreEqual("Guests", pages[1].Security.RoleAssignments[1].Principal);
            Assert.AreEqual("View Only", pages[1].Security.RoleAssignments[1].RoleDefinition);

            Assert.IsTrue(pages[2].Fields.ContainsKey("WikiField"));
            Assert.IsTrue(pages[2].Fields["WikiField"].Contains("The Wiki page HTML code here!"));
            Assert.AreEqual((uint)1, pages[2].WebParts[0].Column);
            Assert.AreEqual((uint)1, pages[2].WebParts[0].Row);
            Assert.AreEqual("Script Editor", pages[2].WebParts[0].Title);
            Assert.IsTrue(pages[2].WebParts[0].Contents.Contains("Microsoft.SharePoint.WebPartPages.ScriptEditorWebPart"));
            Assert.IsTrue(pages[2].WebParts[0].Contents.Contains("showAlert"));
        }

        [TestMethod]
        [TestCategory(TEST_CATEGORY)]
        public void XMLSerializer_Serialize_Pages()
        {
            var provider = new XMLFileSystemTemplateProvider($@"{AppDomain.CurrentDomain.BaseDirectory}\..\..\..\Resources", "Templates");

            var result = new ProvisioningTemplate();

            result.Pages.Add(new PnP.Framework.Provisioning.Model.Page
            {
                Url = "{site}/SitePages/DemoWikiPage.aspx",
                Layout = WikiPageLayout.TwoColumns,
                Overwrite = true
            });

            result.Pages.Add(new PnP.Framework.Provisioning.Model.Page("{site}/SitePages/OneColumnPage.aspx", true, WikiPageLayout.OneColumn, new List<WebPart>(), new PnP.Framework.Provisioning.Model.ObjectSecurity
            {
                ClearSubscopes = true,
                CopyRoleAssignments = true,
                RoleAssignments =
                {
                    new PnP.Framework.Provisioning.Model.RoleAssignment
                    {
                        Principal = "Power Users",
                        RoleDefinition = "Full Control"
                    },
                    new PnP.Framework.Provisioning.Model.RoleAssignment
                    {
                        Principal = "Guests",
                        RoleDefinition = "View Only"
                    }
                }
            }));

            result.Pages.Add(new PnP.Framework.Provisioning.Model.Page("{site}/SitePages/OneColumnPage.aspx", true, WikiPageLayout.OneColumn, new List<WebPart>
            {
                new WebPart
                {
                    Row = 1,
                    Column = 1,
                    Title = "Script Editor",
                    Contents = "<Microsoft.SharePoint.WebPartPages.ScriptEditorWebPart>showAlert</Microsoft.SharePoint.WebPartPages.ScriptEditorWebPart>"
                }
            }, null, new Dictionary<string, string>
            {
                { "WikiField", "The Wiki page HTML code here!"}
            }));


            var serializer = new TargetSerializer();
            provider.SaveAs(result, TEST_OUT_FILE, serializer);

            var path = $"{provider.Connector.Parameters["ConnectionString"]}\\{provider.Connector.Parameters["Container"]}\\{TEST_OUT_FILE}";
            Assert.IsTrue(File.Exists(path)); path = Path.GetFullPath(path);
            var xml = XDocument.Load(path);
            var wrappedResult =
                XMLSerializer.Deserialize<TargetProvisioning>(xml);

            var template = wrappedResult.Templates[0].ProvisioningTemplate.First();
            var pages = template.Pages;

            Assert.AreEqual(TargetNamespace.WikiPageLayout.TwoColumns, pages[0].Layout);
            Assert.AreEqual(true, template.Pages[0].Overwrite);
            Assert.AreEqual("{site}/SitePages/DemoWikiPage.aspx", pages[0].Url);

            Assert.AreEqual(true, pages[1].Security.BreakRoleInheritance.CopyRoleAssignments);
            Assert.AreEqual(true, pages[1].Security.BreakRoleInheritance.ClearSubscopes);
            Assert.AreEqual("Power Users", pages[1].Security.BreakRoleInheritance.RoleAssignment[0].Principal);
            Assert.AreEqual("Full Control", pages[1].Security.BreakRoleInheritance.RoleAssignment[0].RoleDefinition);
            Assert.AreEqual("Guests", pages[1].Security.BreakRoleInheritance.RoleAssignment[1].Principal);
            Assert.AreEqual("View Only", pages[1].Security.BreakRoleInheritance.RoleAssignment[1].RoleDefinition);

            Assert.IsTrue(pages[2].Fields.SingleOrDefault(f => f.FieldName == "WikiField") != null);
            Assert.IsTrue(pages[2].Fields.SingleOrDefault(f => f.FieldName == "WikiField" && f.Value.Contains("The Wiki page HTML code here!")) != null);
            Assert.AreEqual(1, pages[2].WebParts[0].Column);
            Assert.AreEqual(1, pages[2].WebParts[0].Row);
            Assert.AreEqual("Script Editor", pages[2].WebParts[0].Title);
            Assert.IsTrue(pages[2].WebParts[0].Contents.FirstChild.Value.Contains("showAlert"));
        }

        [TestMethod]
        [TestCategory(TEST_CATEGORY)]
        public void XMLSerializer_Deserialize_TermGroups()
        {
            var provider = new XMLFileSystemTemplateProvider($@"{AppDomain.CurrentDomain.BaseDirectory}\..\..\..\Resources", "Templates");

            var serializer = new TargetSerializer();
            var template = provider.GetTemplate(TEST_TEMPLATE, serializer);

            var termGroup = template.TermGroups[0];

            Assert.AreEqual(new Guid("0E8F395E-FF58-4D45-9FF7-E331AB728BEB"), termGroup.Id);
            Assert.AreEqual("{parameter:CompanyName} TermSets", termGroup.Name);
            Assert.AreEqual("user1@contoso.com", termGroup.Contributors[0].Name);
            Assert.AreEqual("user2@contoso.com", termGroup.Managers[0].Name);

            Assert.AreEqual(new Guid("5880B01B-5D6F-4606-A492-3B03A2FB4DD7"), termGroup.TermSets[0].Id);
            Assert.AreEqual(1040, termGroup.TermSets[0].Language);
            Assert.AreEqual("Projects", termGroup.TermSets[0].Name);

            var term = termGroup.TermSets[0].Terms[0];
            Assert.AreEqual("IT Projects", term.Name);

            Assert.AreEqual(new Guid("3D212FC2-F176-4621-AED1-128219666D95"), term.Id);
            Assert.IsTrue(term.Properties.ContainsKey("Property1"));
            Assert.AreEqual("Value1", term.Properties["Property1"]);
            Assert.IsTrue(term.LocalProperties.ContainsKey("LocalProperty1"));
            Assert.AreEqual("Value1", term.LocalProperties["LocalProperty1"]);

            Assert.AreEqual("Cloud", term.Terms[0].Name);
            Assert.AreEqual(new Guid("87C55100-8316-4DA0-97FD-FEB5731880F6"), term.Terms[0].Id);
            Assert.AreEqual("Nuvola", term.Terms[0].Labels[0].Value);
            Assert.AreEqual(1040, term.Terms[0].Labels[0].Language);
            Assert.AreEqual(true, term.Terms[1].IsDeprecated);
        }

        [TestMethod]
        [TestCategory(TEST_CATEGORY)]
        public void XMLSerializer_Serialize_TermGroups()
        {
            var provider = new XMLFileSystemTemplateProvider($@"{AppDomain.CurrentDomain.BaseDirectory}\..\..\..\Resources", "Templates");

            var result = new ProvisioningTemplate();

            result.TermGroups.Add(new PnP.Framework.Provisioning.Model.TermGroup
            {
                Name = "{parameter:CompanyName} TermSets",
                Id = new Guid("0E8F395E-FF58-4D45-9FF7-E331AB728BEB"),
                Contributors =
                {
                    new PnP.Framework.Provisioning.Model.User
                    {
                        Name = "user1@contoso.com"
                    }
                },
                Managers =
                {
                    new PnP.Framework.Provisioning.Model.User
                    {
                        Name = "user2@contoso.com"
                    }
                },
                TermSets =
                {
                    new  PnP.Framework.Provisioning.Model.TermSet
                    {
                        Name = "Projects",
                        Id = new Guid("5880B01B-5D6F-4606-A492-3B03A2FB4DD7"),
                        Language = 1040,
                        Terms =
                        {
                            new PnP.Framework.Provisioning.Model.Term
                            {
                                Name = "IT Projects",
                                Id = new Guid("3D212FC2-F176-4621-AED1-128219666D95"),
                                Properties =
                                {
                                    {"Property1", "Value1" }
                                },
                                LocalProperties =
                                {
                                    {"LocalProperty1", "Value1" }
                                },
                                Terms =
                                {
                                    new PnP.Framework.Provisioning.Model.Term
                                    {
                                        Name = "Cloud",
                                        Id = new Guid("87C55100-8316-4DA0-97FD-FEB5731880F6"),
                                        Labels =
                                        {
                                            new PnP.Framework.Provisioning.Model.TermLabel
                                            {
                                                Value = "Nuvola",
                                                Language = 1040
                                            }
                                        }
                                    },
                                    new PnP.Framework.Provisioning.Model.Term
                                    {
                                        Name = "New Farm",
                                        Id = new Guid("C422BD0D-681D-448F-A41E-C71C473A95CC"),
                                        IsDeprecated = true
                                    }
                                }
                            },
                        }
                    }
                }
            });

            var serializer = new TargetSerializer();
            provider.SaveAs(result, TEST_OUT_FILE, serializer);

            var path = $"{provider.Connector.Parameters["ConnectionString"]}\\{provider.Connector.Parameters["Container"]}\\{TEST_OUT_FILE}";
            Assert.IsTrue(File.Exists(path)); path = Path.GetFullPath(path);
            var xml = XDocument.Load(path);
            var wrappedResult =
                XMLSerializer.Deserialize<TargetProvisioning>(xml);

            var template = wrappedResult.Templates[0].ProvisioningTemplate.First();

            var termGroup = template.TermGroups[0];

            Assert.AreEqual("0E8F395E-FF58-4D45-9FF7-E331AB728BEB".ToLower(), termGroup.ID);
            Assert.AreEqual("{parameter:CompanyName} TermSets", termGroup.Name);
            Assert.AreEqual("user1@contoso.com", termGroup.Contributors[0].Name);
            Assert.AreEqual("user2@contoso.com", termGroup.Managers[0].Name);

            Assert.AreEqual("5880B01B-5D6F-4606-A492-3B03A2FB4DD7".ToLower(), termGroup.TermSets[0].ID);
            Assert.AreEqual(1040, termGroup.TermSets[0].Language);
            Assert.AreEqual("Projects", termGroup.TermSets[0].Name);

            var term = termGroup.TermSets[0].Terms[0];
            Assert.AreEqual("IT Projects", term.Name);

            Assert.AreEqual("3D212FC2-F176-4621-AED1-128219666D95".ToLower(), term.ID);

            Assert.IsTrue(term.CustomProperties.SingleOrDefault(p => p.Key == "Property1") != null);
            Assert.AreEqual("Value1", term.CustomProperties.SingleOrDefault(p => p.Key == "Property1").Value);
            Assert.IsTrue(term.LocalCustomProperties.SingleOrDefault(p => p.Key == "LocalProperty1") != null);
            Assert.AreEqual("Value1", term.LocalCustomProperties.SingleOrDefault(p => p.Key == "LocalProperty1").Value);

            Assert.AreEqual("Cloud", term.Terms.Items[0].Name);
            Assert.AreEqual("87C55100-8316-4DA0-97FD-FEB5731880F6".ToLower(), term.Terms.Items[0].ID);
            Assert.AreEqual("Nuvola", term.Terms.Items[0].Labels[0].Value);
            Assert.AreEqual(1040, term.Terms.Items[0].Labels[0].Language);
            Assert.AreEqual(true, term.Terms.Items[1].IsDeprecated);
        }

        [TestMethod]
        [TestCategory(TEST_CATEGORY)]
        public void XMLSerializer_Deserialize_ComposedLook()
        {
            var provider = new XMLFileSystemTemplateProvider($@"{AppDomain.CurrentDomain.BaseDirectory}\..\..\..\Resources", "Templates");

            var serializer = new TargetSerializer();
            var template = provider.GetTemplate(TEST_TEMPLATE, serializer);

            Assert.IsNotNull(template);
            Assert.IsNotNull(template.ComposedLook);
            Assert.AreEqual("{sitecollection}/Resources/Themes/Contoso/contosobg.jpg", template.ComposedLook.BackgroundFile);
            Assert.AreEqual("{sitecollection}/_catalogs/Theme/15/Custom.spcolor", template.ComposedLook.ColorFile);
            Assert.AreEqual("{sitecollection}/_catalogs/Theme/15/Custom.spfont", template.ComposedLook.FontFile);
            Assert.AreEqual("Custom Look", template.ComposedLook.Name);
            Assert.AreEqual(1, template.ComposedLook.Version);
        }

        [TestMethod]
        [TestCategory(TEST_CATEGORY)]
        public void XMLSerializer_Serialize_ComposedLook()
        {
            var provider = new XMLFileSystemTemplateProvider($@"{AppDomain.CurrentDomain.BaseDirectory}\..\..\..\Resources", "Templates");

            var result = new ProvisioningTemplate();

            var composedLook = new PnP.Framework.Provisioning.Model.ComposedLook()
            {
                BackgroundFile = "{sitecollection}/Resources/Themes/Contoso/contosobg.jpg",
                ColorFile = "{sitecollection}/_catalogs/Theme/15/Custom.spcolor",
                FontFile = "{sitecollection}/_catalogs/Theme/15/Custom.spfont",
                Name = "Custom Look",
                Version = 1
            };
            result.ComposedLook = composedLook;

            var serializer = new TargetSerializer();
            provider.SaveAs(result, TEST_OUT_FILE, serializer);

            var path = $"{provider.Connector.Parameters["ConnectionString"]}\\{provider.Connector.Parameters["Container"]}\\{TEST_OUT_FILE}";
            Assert.IsTrue(File.Exists(path)); path = Path.GetFullPath(path);
            var xml = XDocument.Load(path);
            TargetProvisioning wrappedResult =
                XMLSerializer.Deserialize<TargetProvisioning>(xml);

            var template = wrappedResult.Templates[0].ProvisioningTemplate.First();

            Assert.IsNotNull(template);
            Assert.IsNotNull(template.ComposedLook);
            Assert.AreEqual("{sitecollection}/Resources/Themes/Contoso/contosobg.jpg", template.ComposedLook.BackgroundFile);
            Assert.AreEqual("{sitecollection}/_catalogs/Theme/15/Custom.spcolor", template.ComposedLook.ColorFile);
            Assert.AreEqual("{sitecollection}/_catalogs/Theme/15/Custom.spfont", template.ComposedLook.FontFile);
            Assert.AreEqual("Custom Look", template.ComposedLook.Name);
            Assert.AreEqual(1, template.ComposedLook.Version);
        }

        [TestMethod]
        [TestCategory(TEST_CATEGORY)]
        public void XMLSerializer_Deserialize_SearchSettings()
        {
            var provider = new XMLFileSystemTemplateProvider($@"{AppDomain.CurrentDomain.BaseDirectory}\..\..\..\Resources", "Templates");

            var serializer = new TargetSerializer();
            var template = provider.GetTemplate(TEST_TEMPLATE, serializer);

            Assert.IsNotNull(template.SiteSearchSettings);
            Assert.IsTrue(template.SiteSearchSettings.Contains("SearchQueryConfigurationSettings"));
            Assert.IsTrue(template.SiteSearchSettings.Contains("BestBets"));
            Assert.IsTrue(template.SiteSearchSettings.Contains("SearchRankingModelConfigurationSettings"));
            Assert.IsTrue(template.SiteSearchSettings.Contains("ManagedProperties"));
            Assert.IsTrue(template.SiteSearchSettings.Contains("CrawledProperties"));
            Assert.IsTrue(template.SiteSearchSettings.Contains("Mappings"));
            Assert.IsTrue(template.SiteSearchSettings.Contains("Overrides"));
        }

        [TestMethod]
        [TestCategory(TEST_CATEGORY)]
        public void XMLSerializer_Serialize_SearchSettings()
        {
            var provider = new XMLFileSystemTemplateProvider($@"{AppDomain.CurrentDomain.BaseDirectory}\..\..\..\Resources", "Templates");

            var result = new ProvisioningTemplate
            {
                SiteSearchSettings = "<SearchConfigurationSettings></SearchConfigurationSettings>"
            };

            var serializer = new TargetSerializer();
            provider.SaveAs(result, TEST_OUT_FILE, serializer);

            var path = $"{provider.Connector.Parameters["ConnectionString"]}\\{provider.Connector.Parameters["Container"]}\\{TEST_OUT_FILE}";
            Assert.IsTrue(File.Exists(path)); path = Path.GetFullPath(path);
            var xml = XDocument.Load(path);
            TargetProvisioning wrappedResult =
                XMLSerializer.Deserialize<TargetProvisioning>(xml);

            var template = wrappedResult.Templates[0].ProvisioningTemplate.First();

            Assert.AreEqual("SearchConfigurationSettings", template.SearchSettings.SiteSearchSettings.Name);
        }

        [TestMethod]
        [TestCategory(TEST_CATEGORY)]
        public void XMLSerializer_Deserialize_Publishing()
        {
            var provider = new XMLFileSystemTemplateProvider($@"{AppDomain.CurrentDomain.BaseDirectory}\..\..\..\Resources", "Templates");

            var serializer = new TargetSerializer();
            var template = provider.GetTemplate(TEST_TEMPLATE, serializer);

            var publishing = template.Publishing;

            Assert.AreEqual(AutoCheckRequirementsOptions.MakeCompliant, publishing.AutoCheckRequirements);
            Assert.AreEqual("CustomDesign.wsp", publishing.DesignPackage.DesignPackagePath);
            Assert.AreEqual(1, publishing.DesignPackage.MajorVersion);
            Assert.AreEqual(0, publishing.DesignPackage.MinorVersion);
            Assert.AreEqual(new Guid("A3349210-5283-44A5-A23F-00F489EB690B"), publishing.DesignPackage.PackageGuid);
            Assert.AreEqual("Custom Design", publishing.DesignPackage.PackageName);

            Assert.AreEqual(1033, publishing.AvailableWebTemplates[0].LanguageCode);
            Assert.AreEqual("STS#0", publishing.AvailableWebTemplates[0].TemplateName);
            Assert.AreEqual("News.aspx", publishing.PageLayouts[0].Path);
            Assert.AreEqual(true, publishing.PageLayouts[1].IsDefault);

            Assert.AreEqual(100, publishing.ImageRenditions[0].Width);
            Assert.AreEqual(100, publishing.ImageRenditions[0].Height);
            Assert.AreEqual("SmallSquare", publishing.ImageRenditions[0].Name);
        }

        [TestMethod]
        [TestCategory(TEST_CATEGORY)]
        public void XMLSerializer_Serialize_Publishing()
        {
            var provider = new XMLFileSystemTemplateProvider($@"{AppDomain.CurrentDomain.BaseDirectory}\..\..\..\Resources", "Templates");

            var result = new ProvisioningTemplate
            {
                Publishing = new PnP.Framework.Provisioning.Model.Publishing
                {
                    AutoCheckRequirements = AutoCheckRequirementsOptions.MakeCompliant,
                    DesignPackage = new PnP.Framework.Provisioning.Model.DesignPackage
                    {
                        DesignPackagePath = "CustomDesign.wsp",
                        MajorVersion = 1,
                        MinorVersion = 0,
                        PackageGuid = new Guid("A3349210-5283-44A5-A23F-00F489EB690B"),
                        PackageName = "Custom Design"
                    },
                    AvailableWebTemplates =
                {
                    new PnP.Framework.Provisioning.Model.AvailableWebTemplate
                    {
                        LanguageCode = 1033,
                        TemplateName = "STS#0"
                    }
                },
                    PageLayouts =
                {
                    new PnP.Framework.Provisioning.Model.PageLayout
                    {
                        Path = "News.aspx"
                    },
                    new PnP.Framework.Provisioning.Model.PageLayout
                    {
                        Path = "SimplePage.aspx",
                        IsDefault = true
                    }
                },
                    ImageRenditions =
                {
                    new PnP.Framework.Provisioning.Model.ImageRendition
                    {
                        Name = "SmallSquare",
                        Height = 100,
                        Width = 100
                    }
                }
                }
            };

            var serializer = new TargetSerializer();
            provider.SaveAs(result, TEST_OUT_FILE, serializer);

            var path = $"{provider.Connector.Parameters["ConnectionString"]}\\{provider.Connector.Parameters["Container"]}\\{TEST_OUT_FILE}";
            Assert.IsTrue(File.Exists(path)); path = Path.GetFullPath(path);
            var xml = XDocument.Load(path);
            TargetProvisioning wrappedResult =
                XMLSerializer.Deserialize<TargetProvisioning>(xml);

            var template = wrappedResult.Templates[0].ProvisioningTemplate.First();
            var publishing = template.Publishing;

            Assert.AreEqual(PublishingAutoCheckRequirements.MakeCompliant, publishing.AutoCheckRequirements);
            Assert.AreEqual("CustomDesign.wsp", publishing.DesignPackage.DesignPackagePath);
            Assert.AreEqual(1, publishing.DesignPackage.MajorVersion);
            Assert.AreEqual(0, publishing.DesignPackage.MinorVersion);
            Assert.AreEqual("a3349210-5283-44a5-a23f-00f489eb690b", publishing.DesignPackage.PackageGuid);
            Assert.AreEqual("Custom Design", publishing.DesignPackage.PackageName);

            Assert.AreEqual(1033, publishing.AvailableWebTemplates[0].LanguageCode);
            Assert.AreEqual("STS#0", publishing.AvailableWebTemplates[0].TemplateName);
            Assert.AreEqual("News.aspx", publishing.PageLayouts.PageLayout[0].Path);
            Assert.AreEqual("SimplePage.aspx", publishing.PageLayouts.Default);

            Assert.AreEqual("100", publishing.ImageRenditions[0].Width);
            Assert.AreEqual("100", publishing.ImageRenditions[0].Height);
            Assert.AreEqual("SmallSquare", publishing.ImageRenditions[0].Name);
        }

        [TestMethod]
        [TestCategory(TEST_CATEGORY)]
        public void XMLSerializer_Deserialize_SiteWebhooks()
        {
            var provider = new XMLFileSystemTemplateProvider($@"{AppDomain.CurrentDomain.BaseDirectory}\..\..\..\Resources", "Templates");

            var serializer = new TargetSerializer();
            var template = provider.GetTemplate(TEST_TEMPLATE, serializer);

            var webhooks = template.SiteWebhooks;

            Assert.AreEqual(120, webhooks[0].ExpiresInDays);
            Assert.AreEqual(SiteWebhookType.WebCreated, webhooks[0].SiteWebhookType);
            Assert.AreEqual("http://myapp.azurewebsites.net/WebHookListener", webhooks[0].ServerNotificationUrl);
        }

        [TestMethod]
        [TestCategory(TEST_CATEGORY)]
        public void XMLSerializer_Serialize_SiteWebhooks()
        {
            var provider = new XMLFileSystemTemplateProvider($@"{AppDomain.CurrentDomain.BaseDirectory}\..\..\..\Resources", "Templates");

            var result = new ProvisioningTemplate();

            result.SiteWebhooks.Add(new PnP.Framework.Provisioning.Model.SiteWebhook
            {
                SiteWebhookType = SiteWebhookType.WebCreated,
                ServerNotificationUrl = "http://myapp.azurewebsites.net/WebHookListener",
                ExpiresInDays = 120
            });

            var serializer = new TargetSerializer();
            provider.SaveAs(result, TEST_OUT_FILE, serializer);

            var path = $"{provider.Connector.Parameters["ConnectionString"]}\\{provider.Connector.Parameters["Container"]}\\{TEST_OUT_FILE}";
            Assert.IsTrue(File.Exists(path)); path = Path.GetFullPath(path);
            var xml = XDocument.Load(path);
            TargetProvisioning wrappedResult =
                XMLSerializer.Deserialize<TargetProvisioning>(xml);

            var template = wrappedResult.Templates[0].ProvisioningTemplate.First();
            var publishing = template.Publishing;

            var webhooks = template.SiteWebhooks;

            Assert.AreEqual("120", webhooks[0].ExpiresInDays);
            Assert.AreEqual(SiteWebhookSiteWebhookType.WebCreated, webhooks[0].SiteWebhookType);
            Assert.AreEqual("http://myapp.azurewebsites.net/WebHookListener", webhooks[0].ServerNotificationUrl);
        }

        [TestMethod]
        [TestCategory(TEST_CATEGORY)]
        public void XMLSerializer_Deserialize_ClientSidePages()
        {
            var provider = new XMLFileSystemTemplateProvider($@"{AppDomain.CurrentDomain.BaseDirectory}\..\..\..\Resources", "Templates");

            var serializer = new TargetSerializer();
            var template = provider.GetTemplate(TEST_TEMPLATE, serializer);

            var clientSidePages = template.ClientSidePages;

            Assert.AreEqual("SamplePage", clientSidePages[0].PageName);
            Assert.AreEqual(true, clientSidePages[0].PromoteAsTemplate);
            Assert.AreEqual(true, clientSidePages[0].PromoteAsNewsArticle);
            Assert.AreEqual(true, clientSidePages[0].Overwrite);
            Assert.AreEqual(true, clientSidePages[0].Publish);
            Assert.AreEqual("Article", clientSidePages[0].Layout);
            Assert.AreEqual(true, clientSidePages[0].EnableComments);
            Assert.AreEqual("Client Side Page Title", clientSidePages[0].Title);
            Assert.AreEqual("0x01010012345", clientSidePages[0].ContentTypeID);
            Assert.AreEqual(true, clientSidePages[0].CreateTranslations);
            Assert.AreEqual("images/pageThumbnail.png", clientSidePages[0].ThumbnailUrl);

            var page = clientSidePages[0];
            // header
            Assert.AreEqual(PnP.Framework.Provisioning.Model.ClientSidePageHeaderType.Custom, page.Header.Type);
            Assert.AreEqual("./site%20assets/picture.png", page.Header.ServerRelativeImageUrl);
            Assert.AreEqual(10.56, page.Header.TranslateX);
            Assert.AreEqual(15.12345, page.Header.TranslateY);
            Assert.AreEqual(PnP.Framework.Provisioning.Model.ClientSidePageHeaderLayoutType.FullWidthImage, page.Header.LayoutType);
            Assert.AreEqual(PnP.Framework.Provisioning.Model.ClientSidePageHeaderTextAlignment.Center, page.Header.TextAlignment);
            Assert.AreEqual("Alternate text", page.Header.AlternativeText);
            Assert.AreEqual("John Black, Mike White", page.Header.Authors);
            Assert.AreEqual("Bill Green", page.Header.AuthorByLine);
            Assert.AreEqual(5, page.Header.AuthorByLineId);
            Assert.AreEqual(true, page.Header.ShowPublishDate);
            Assert.AreEqual(true, page.Header.ShowTopicHeader);
            Assert.AreEqual("Topic header value", page.Header.TopicHeader);
            Assert.IsTrue(page.Header.ShowBackgroundGradient);

            var section = page.Sections[0];

            // sections
            Assert.AreEqual(1, section.Order);
            Assert.AreEqual(PnP.Framework.Provisioning.Model.CanvasSectionType.OneColumn, section.Type);
            Assert.AreEqual(PnP.Framework.Provisioning.Model.Emphasis.Neutral, section.BackgroundEmphasis);
            Assert.IsTrue(section.Collapsible);
            Assert.AreEqual("My collapsible section", section.DisplayName);
            Assert.AreEqual(IconAlignment.Left, section.IconAlignment);
            Assert.IsTrue(section.IsExpanded);
            Assert.IsTrue(section.ShowDividerLine);
            Assert.AreEqual(PnP.Framework.Provisioning.Model.Emphasis.Neutral, section.VerticalSectionEmphasis);

            Assert.AreEqual("...", section.Controls[0].CustomWebPartName);
            Assert.AreEqual(WebPartType.Image, section.Controls[0].Type);
            Assert.AreEqual("{}", section.Controls[0].JsonControlData);
            Assert.AreEqual(new Guid("0eaba53f-55d8-44b5-9f7c-61301c7f1e0e"), section.Controls[0].ControlId);
            Assert.AreEqual(1, section.Controls[0].Order);
            Assert.IsTrue(section.Controls[0].ControlProperties.ContainsKey("Key1"));
            Assert.AreEqual("{token}", section.Controls[0].ControlProperties["Key1"]);

            // field values
            Assert.IsTrue(page.FieldValues.ContainsKey("Category"));
            Assert.AreEqual("Marketing", page.FieldValues["Category"]);

            // properties
            Assert.IsTrue(page.Properties.ContainsKey("Key01"));
            Assert.AreEqual("Value 01", page.Properties["Key01"]);

            // security
            Assert.AreEqual(true, page.Security.ClearSubscopes);
            Assert.AreEqual(false, page.Security.CopyRoleAssignments);
            Assert.AreEqual("user1@contoso.com", page.Security.RoleAssignments[0].Principal);
            Assert.AreEqual("Full Control", page.Security.RoleAssignments[0].RoleDefinition);
        }

        [TestMethod]
        [TestCategory(TEST_CATEGORY)]
        public void XMLSerializer_Serialize_ClientSidePages()
        {
            var provider = new XMLFileSystemTemplateProvider($@"{AppDomain.CurrentDomain.BaseDirectory}\..\..\..\Resources", "Templates");

            var result = new ProvisioningTemplate();

            result.ClientSidePages.Add(new PnP.Framework.Provisioning.Model.ClientSidePage
            {
                PageName = "SamplePage",
                PromoteAsTemplate = true,
                PromoteAsNewsArticle = true,
                Overwrite = true,
                Publish = true,
                Layout = "Article",
                EnableComments = true,
                Title = "Client Side Page Title",
                ContentTypeID = "0x01010012345",
                CreateTranslations = true,
                ThumbnailUrl = "images/pageThumbnail.png",
                LCID = 1040,
                Header = new PnP.Framework.Provisioning.Model.ClientSidePageHeader
                {
                    Type = ClientSidePageHeaderType.Custom,
                    ServerRelativeImageUrl = "./site%20assets/picture.png",
                    TranslateX = 10.56,
                    TranslateY = 15.12345,
                    LayoutType = ClientSidePageHeaderLayoutType.FullWidthImage,
                    TextAlignment = ClientSidePageHeaderTextAlignment.Center,
                    AlternativeText = "Alternate text",
                    Authors = "John Black, Mike White",
                    AuthorByLine = "Bill Green",
                    AuthorByLineId = 5,
                    ShowPublishDate = true,
                    ShowTopicHeader = true,
                    TopicHeader = "Topic header value",
                    ShowBackgroundGradient = true,
                },
                Sections =
                {
                    new PnP.Framework.Provisioning.Model.CanvasSection
                    {
                        Order = 1,
                        Type = CanvasSectionType.OneColumnVerticalSection,
                        BackgroundEmphasis = PnP.Framework.Provisioning.Model.Emphasis.Soft,
                        Collapsible = true,
                        DisplayName = "My collapsible section",
                        IconAlignment = IconAlignment.Left,
                        IsExpanded = true,
                        ShowDividerLine = true,
                        VerticalSectionEmphasis = PnP.Framework.Provisioning.Model.Emphasis.Neutral,
                        Controls =
                        {
                            new PnP.Framework.Provisioning.Model.CanvasControl
                            {
                                CustomWebPartName = "...",
                                Type = WebPartType.Image,
                                JsonControlData = "{}",
                                ControlId = new Guid("0eaba53f-55d8-44b5-9f7c-61301c7f1e0e"),
                                Order = 1,
                                ControlProperties =
                                {
                                    {"Key1", "{token}" }
                                }
                            }
                        }
                    }
                },
                FieldValues =
                {
                    { "Category","Marketing" }
                },
                Properties =
                {
                    { "Key01", "Value 01" }
                },
            });

            result.ClientSidePages[0].Security.ClearSubscopes = true;
            result.ClientSidePages[0].Security.CopyRoleAssignments = false;
            result.ClientSidePages[0].Security.RoleAssignments.Add(new PnP.Framework.Provisioning.Model.RoleAssignment
            {
                Principal = "user1@contoso.com",
                RoleDefinition = "Full Control"
            });

            var serializer = new TargetSerializer();
            provider.SaveAs(result, TEST_OUT_FILE, serializer);

            var path = $"{provider.Connector.Parameters["ConnectionString"]}\\{provider.Connector.Parameters["Container"]}\\{TEST_OUT_FILE}";
            Assert.IsTrue(File.Exists(path)); path = Path.GetFullPath(path);
            var xml = XDocument.Load(path);
            TargetProvisioning wrappedResult =
                XMLSerializer.Deserialize<TargetProvisioning>(xml);

            var template = wrappedResult.Templates[0].ProvisioningTemplate.First();
            var clientSidePages = template.ClientSidePages;

            Assert.AreEqual("SamplePage", clientSidePages[0].PageName);
            Assert.AreEqual(true, clientSidePages[0].PromoteAsTemplate);
            Assert.AreEqual(true, clientSidePages[0].PromoteAsNewsArticle);
            Assert.AreEqual(true, clientSidePages[0].Overwrite);
            Assert.AreEqual(true, clientSidePages[0].Publish);
            Assert.AreEqual("Article", clientSidePages[0].Layout);
            Assert.AreEqual(true, clientSidePages[0].EnableComments);
            Assert.AreEqual("Client Side Page Title", clientSidePages[0].Title);
            Assert.AreEqual("0x01010012345", clientSidePages[0].ContentTypeID);
            Assert.AreEqual(true, clientSidePages[0].CreateTranslations);
            Assert.AreEqual("images/pageThumbnail.png", clientSidePages[0].ThumbnailUrl);
            Assert.AreEqual(1040, clientSidePages[0].LCID);

            var page = clientSidePages[0];
            // header
            Assert.AreEqual(TargetNamespace.BaseClientSidePageHeaderType.Custom, page.Header.Type);
            Assert.AreEqual("./site%20assets/picture.png", page.Header.ServerRelativeImageUrl);
            Assert.AreEqual(10.56, page.Header.TranslateX);
            Assert.AreEqual(15.12345, page.Header.TranslateY);
            Assert.AreEqual(TargetNamespace.BaseClientSidePageHeaderLayoutType.FullWidthImage, page.Header.LayoutType);
            Assert.AreEqual(TargetNamespace.BaseClientSidePageHeaderTextAlignment.Center, page.Header.TextAlignment);
            Assert.AreEqual("Alternate text", page.Header.AlternativeText);
            Assert.AreEqual("John Black, Mike White", page.Header.Authors);
            Assert.AreEqual("Bill Green", page.Header.AuthorByLine);
            Assert.AreEqual(5, page.Header.AuthorByLineId);
            Assert.AreEqual(true, page.Header.ShowPublishDate);
            Assert.AreEqual(true, page.Header.ShowTopicHeader);
            Assert.AreEqual("Topic header value", page.Header.TopicHeader);
            Assert.IsTrue(page.Header.ShowBackgroundGradient);
            Assert.IsTrue(page.Header.ShowBackgroundGradientSpecified);

            var section = page.Sections[0];

            // sections
            Assert.AreEqual(1, section.Order);
            Assert.AreEqual(TargetNamespace.CanvasSectionType.OneColumnVerticalSection, section.Type);
            Assert.AreEqual(TargetNamespace.Emphasis.Soft, section.BackgroundEmphasis);
            Assert.IsTrue(section.Collapsible);
            Assert.AreEqual("My collapsible section", section.DisplayName);
            Assert.AreEqual(TargetNamespace.CanvasSectionIconAlignment.Left, section.IconAlignment);
            Assert.IsTrue(section.IsExpanded);
            Assert.IsTrue(section.ShowDividerLine);
            Assert.AreEqual(TargetNamespace.Emphasis.Neutral, section.VerticalSectionEmphasis);

            Assert.AreEqual("...", section.Controls[0].CustomWebPartName);
            Assert.AreEqual(TargetNamespace.CanvasControlWebPartType.Image, section.Controls[0].WebPartType);
            Assert.AreEqual("{}", section.Controls[0].JsonControlData);
            Assert.AreEqual("0eaba53f-55d8-44b5-9f7c-61301c7f1e0e", section.Controls[0].ControlId);
            Assert.AreEqual(1, section.Controls[0].Order);
            Assert.IsTrue(section.Controls[0].CanvasControlProperties.SingleOrDefault(p => p.Key == "Key1") != null);
            Assert.AreEqual("{token}", section.Controls[0].CanvasControlProperties.SingleOrDefault(p => p.Key == "Key1").Value);

            // field values
            Assert.IsTrue(page.FieldValues.SingleOrDefault(p => p.Key == "Category") != null);
            Assert.AreEqual("Marketing", page.FieldValues.SingleOrDefault(p => p.Key == "Category").Value);

            // properties
            Assert.IsTrue(page.Properties.SingleOrDefault(p => p.Key == "Key01") != null);
            Assert.AreEqual("Value 01", page.Properties.SingleOrDefault(p => p.Key == "Key01").Value);

            // security
            Assert.AreEqual(true, page.Security.BreakRoleInheritance.ClearSubscopes);
            Assert.AreEqual(false, page.Security.BreakRoleInheritance.CopyRoleAssignments);
            Assert.AreEqual("user1@contoso.com", page.Security.BreakRoleInheritance.RoleAssignment[0].Principal);
            Assert.AreEqual("Full Control", page.Security.BreakRoleInheritance.RoleAssignment[0].RoleDefinition);
        }
    }
}
