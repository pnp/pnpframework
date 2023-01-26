﻿using Microsoft.Online.SharePoint.TenantAdministration;
using Microsoft.SharePoint.Client;
using PnP.Core.Services;
using PnP.Framework.Diagnostics;
using PnP.Framework.Provisioning.Model;
using PnP.Framework.Provisioning.ObjectHandlers.Utilities;
using PnP.Framework.Utilities;
using System;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Web;

namespace PnP.Framework.Provisioning.ObjectHandlers
{
    internal class ObjectWebSettings : ObjectHandlerBase
    {
        public override string Name
        {
            get { return "Web Settings"; }
        }

        public override string InternalName => "WebSettings";
        public override ProvisioningTemplate ExtractObjects(Web web, ProvisioningTemplate template, ProvisioningTemplateCreationInformation creationInfo)
        {
            using (var scope = new PnPMonitoredScope(this.Name))
            {
                web.EnsureProperties(
                    w => w.NoCrawl,
                    w => w.CommentsOnSitePagesDisabled,
                    w => w.ExcludeFromOfflineClient,
                    w => w.MembersCanShare,
                    w => w.DisableFlows,
                    w => w.DisableAppViews,
                    w => w.HorizontalQuickLaunch,
                    w => w.QuickLaunchEnabled,
                    w => w.SearchScope,
                    w => w.SearchBoxInNavBar,
                    //w => w.Title,
                    //w => w.Description,
                    w => w.MasterUrl,
                    w => w.CustomMasterUrl,
                    w => w.SiteLogoUrl,
                    w => w.RequestAccessEmail,
                    w => w.RootFolder,
                    w => w.AlternateCssUrl,
                    w => w.ServerRelativeUrl,
                    w => w.Url
                    );

                var webSettings = new WebSettings
                {
                    NoCrawl = web.NoCrawl,
                    CommentsOnSitePagesDisabled = web.CommentsOnSitePagesDisabled,
                    ExcludeFromOfflineClient = web.ExcludeFromOfflineClient,
                    MembersCanShare = web.MembersCanShare,
                    DisableFlows = web.DisableFlows,
                    DisableAppViews = web.DisableAppViews,
                    HorizontalQuickLaunch = web.HorizontalQuickLaunch,
                    QuickLaunchEnabled = web.QuickLaunchEnabled,
                    SearchScope = (SearchScopes)Enum.Parse(typeof(SearchScopes), web.SearchScope.ToString(), true),
                    SearchBoxInNavBar = (SearchBoxInNavBar)Enum.Parse(typeof(SearchBoxInNavBar), web.SearchBoxInNavBar.ToString(), true),
                    SearchCenterUrl = web.GetWebSearchCenterUrl(true),
                    // We're not extracting Title and Description
                    //webSettings.Title = Tokenize(web.Title, web.Url);
                    //webSettings.Description = Tokenize(web.Description, web.Url);
                    MasterPageUrl = Tokenize(web.MasterUrl, web.Url),
                    CustomMasterPageUrl = Tokenize(web.CustomMasterUrl, web.Url),
                    SiteLogo = TokenizeHost(web, Tokenize(web.SiteLogoUrl, web.Url)),
                    // Notice. No tokenization needed for the welcome page, it's always relative for the site
                    WelcomePage = web.RootFolder.WelcomePage,
                    AlternateCSS = Tokenize(web.AlternateCssUrl, web.Url),
                    RequestAccessEmail = web.RequestAccessEmail
                };

                // Can we get the hubsite url? This requires Tenant Admin rights
                try
                {
                    var site = ((ClientContext)web.Context).Site;
                    site.EnsureProperties(s => s.HubSiteId, s => s.Id);
                    if (site.HubSiteId != Guid.Empty && site.HubSiteId != site.Id)
                    {
                        if (TenantExtensions.IsCurrentUserTenantAdmin(web.Context as ClientContext))
                        {
                            using (var tenantContext = web.Context.Clone((web.Context as ClientContext).Web.GetTenantAdministrationUrl()))
                            {
                                var tenant = new Tenant(tenantContext);
                                var hubsiteProperties = tenant.GetHubSitePropertiesById(site.HubSiteId);
                                tenantContext.Load(hubsiteProperties);
                                tenantContext.ExecuteQueryRetry();
                                webSettings.HubSiteUrl = TokenizeHost(web,hubsiteProperties.SiteUrl);
                            }
                        }
                        else
                        {
                            WriteMessage("You need to be a SharePoint admin to extract Hub site Url.", ProvisioningMessageType.Warning);
                        }
                    }
                }
                catch { }

                if (creationInfo.PersistBrandingFiles)
                {
                    if (!string.IsNullOrEmpty(web.MasterUrl))
                    {
                        var masterUrl = web.MasterUrl.ToLower();
                        if (!masterUrl.EndsWith("default.master") && !masterUrl.EndsWith("custom.master") && !masterUrl.EndsWith("v4.master") && !masterUrl.EndsWith("seattle.master") && !masterUrl.EndsWith("oslo.master"))
                        {

                            if (PersistFile(web, creationInfo, scope, web.MasterUrl))
                            {
                                template.Files.Add(GetTemplateFile(web, web.MasterUrl));
                            }
                        }
                    }
                    if (!string.IsNullOrEmpty(web.CustomMasterUrl))
                    {
                        var customMasterUrl = web.CustomMasterUrl.ToLower();
                        if (!customMasterUrl.EndsWith("default.master") && !customMasterUrl.EndsWith("custom.master") && !customMasterUrl.EndsWith("v4.master") && !customMasterUrl.EndsWith("seattle.master") && !customMasterUrl.EndsWith("oslo.master"))
                        {
                            if (PersistFile(web, creationInfo, scope, web.CustomMasterUrl))
                            {
                                template.Files.Add(GetTemplateFile(web, web.CustomMasterUrl));
                            }
                        }
                    }
                    // Extract site logo if property has been set and it's not dynamic image from _api URL
                    if (!string.IsNullOrEmpty(web.SiteLogoUrl) && (!web.SiteLogoUrl.ToLowerInvariant().Contains("_api/")))
                    {
                        // Convert to server relative URL if needed (web.SiteLogoUrl can be set to FQDN URL of a file hosted in the site (e.g. for communication sites))
                        Uri webUri = new Uri(web.Url);
                        string webUrl = $"{webUri.Scheme}://{webUri.DnsSafeHost}";
                        string siteLogoServerRelativeUrl = web.SiteLogoUrl.Replace(webUrl, "");

                        if (PersistFile(web, creationInfo, scope, siteLogoServerRelativeUrl))
                        {
                            template.Files.Add(GetTemplateFile(web, Uri.UnescapeDataString(siteLogoServerRelativeUrl)));
                        }
                    }
                    if (!string.IsNullOrEmpty(web.AlternateCssUrl))
                    {
                        if (PersistFile(web, creationInfo, scope, web.AlternateCssUrl))
                        {
                            template.Files.Add(GetTemplateFile(web, web.AlternateCssUrl));
                        }
                    }
                    var files = template.Files.Distinct().ToList();
                    template.Files.Clear();
                    template.Files.AddRange(files);
                }

                if (!creationInfo.PersistBrandingFiles)
                {
                    if (creationInfo.BaseTemplate != null)
                    {
                        if (!webSettings.Equals(creationInfo.BaseTemplate.WebSettings))
                        {
                            template.WebSettings = webSettings;
                        }
                    }
                    else
                    {
                        template.WebSettings = webSettings;
                    }
                }
                else
                {
                    template.WebSettings = webSettings;
                }
            }
            return template;
        }

        //TODO: Candidate for cleanup
        private Model.File GetTemplateFile(Web web, string serverRelativeUrl)
        {

            var webServerUrl = web.EnsureProperty(w => w.Url);
            var serverUri = new Uri(webServerUrl);
            var serverUrl = $"{serverUri.Scheme}://{serverUri.Authority}";
            var fullUri = new Uri(UrlUtility.Combine(serverUrl, serverRelativeUrl));

            var folderPath = Uri.UnescapeDataString(fullUri.Segments.Take(fullUri.Segments.Length - 1).ToArray().Aggregate((i, x) => i + x).TrimEnd('/'));
            var fileName = fullUri.Segments[fullUri.Segments.Length - 1];

            // store as site relative path
            folderPath = folderPath.Replace(web.ServerRelativeUrl, "").Trim('/');
            var templateFile = new Model.File()
            {
                Folder = Tokenize(folderPath, web.Url),
                Src = !string.IsNullOrEmpty(folderPath) ? $"{folderPath}/{fileName}" : fileName,
                Overwrite = true,
            };

            return templateFile;
        }

        private bool PersistFile(Web web, ProvisioningTemplateCreationInformation creationInfo, PnPMonitoredScope scope, string serverRelativeUrl)
        {
            var success = false;
            if (creationInfo.PersistBrandingFiles)
            {
                if (creationInfo.FileConnector != null)
                {
                    if (UrlUtility.IsIisVirtualDirectory(serverRelativeUrl))
                    {
                        scope.LogWarning("File is not located in the content database. Not retrieving {0}", serverRelativeUrl);
                        return success;
                    }

                    try
                    {
                        var file = web.GetFileByServerRelativePath(ResourcePath.FromDecodedUrl(serverRelativeUrl));
                        string fileName = string.Empty;
                        if (serverRelativeUrl.IndexOf("/") > -1)
                        {
                            fileName = serverRelativeUrl.Substring(serverRelativeUrl.LastIndexOf("/") + 1);
                        }
                        else
                        {
                            fileName = serverRelativeUrl;
                        }
                        web.Context.Load(file, f => f.ServerRelativePath);
                        web.Context.ExecuteQueryRetry();
                        ClientResult<Stream> stream = file.OpenBinaryStream();
                        web.Context.ExecuteQueryRetry();

                        var baseUri = new Uri(web.Url);
                        var fullUri = new Uri(baseUri, file.ServerRelativePath.DecodedUrl);
                        var folderPath = Uri.UnescapeDataString(fullUri.Segments.Take(fullUri.Segments.Length - 1).ToArray().Aggregate((i, x) => i + x).TrimEnd('/'));

                        // Configure the filename to use 
                        fileName = Uri.UnescapeDataString(fullUri.Segments[fullUri.Segments.Length - 1]);

                        // Build up a site relative container URL...might end up empty as well
                        String container = Uri.UnescapeDataString(folderPath.Replace(web.ServerRelativeUrl, "")).Trim('/').Replace("/", "\\");

                        using (Stream memStream = new MemoryStream())
                        {
                            CopyStream(stream.Value, memStream);
                            memStream.Position = 0;
                            if (!string.IsNullOrEmpty(container))
                            {
                                creationInfo.FileConnector.SaveFileStream(fileName, container, memStream);
                            }
                            else
                            {
                                creationInfo.FileConnector.SaveFileStream(fileName, memStream);
                            }
                        }
                        success = true;
                    }
                    catch (ServerException ex1)
                    {
                        // If we are referring a file from a location outside of the current web or at a location where we cannot retrieve the file an exception is thrown. We swallow this exception.
                        if (ex1.ServerErrorCode != -2147024809)
                        {
                            throw;
                        }
                        else
                        {
                            scope.LogWarning("File is not necessarily located in the current web. Not retrieving {0}", serverRelativeUrl);
                        }
                    }
                }
                else
                {
                    WriteMessage("No connector present to persist site logo.", ProvisioningMessageType.Error);
                    scope.LogError("No connector present to persist site logo");
                }
            }
            else
            {
                success = true;
            }
            return success;
        }

        private static string TokenizeHost(Web web, string json)
        {
            // HostUrl token replacement
            var uri = new Uri(web.Url);
            json = Regex.Replace(json, $"{uri.Scheme}://{uri.DnsSafeHost}:{uri.Port}", "{hosturl}", RegexOptions.IgnoreCase);
            json = Regex.Replace(json, $"{uri.Scheme}://{uri.DnsSafeHost}", "{hosturl}", RegexOptions.IgnoreCase);
            return json;
        }


        private static void CopyStream(Stream source, Stream destination)
        {
            byte[] buffer = new byte[32768];
            int bytesRead;

            do
            {
                bytesRead = source.Read(buffer, 0, buffer.Length);
                destination.Write(buffer, 0, bytesRead);
            } while (bytesRead != 0);
        }

        public override TokenParser ProvisionObjects(Web web, ProvisioningTemplate template, TokenParser parser, ProvisioningTemplateApplyingInformation applyingInformation)
        {
            using (var scope = new PnPMonitoredScope(this.Name))
            {
                if (template.WebSettings != null)
                {
                    // Check if this is not a noscript site as we're not allowed to update some properties
                    bool isNoScriptSite = web.IsNoScriptSite();

                    web.EnsureProperties(
                        w => w.NoCrawl,
                        w => w.CommentsOnSitePagesDisabled,
                        w => w.ExcludeFromOfflineClient,
                        w => w.MembersCanShare,
                        w => w.DisableFlows,
                        w => w.DisableAppViews,
                        w => w.HorizontalQuickLaunch,
                        w => w.SearchScope,
                        w => w.SearchBoxInNavBar,
                        w => w.RootFolder,
                        w => w.Title,
                        w => w.Description,
                        w => w.AlternateCssUrl,
                        w => w.WebTemplate,
                        w => w.HasUniqueRoleAssignments);

                    var webSettings = template.WebSettings;

                    // Since the IsSubSite function can trigger an executequery ensure it's called before any updates to the web object are done.
                    if (!web.IsSubSite() || (web.IsSubSite() && web.HasUniqueRoleAssignments))
                    {
                        String requestAccessEmailValue = parser.ParseString(webSettings.RequestAccessEmail);
                        if (!String.IsNullOrEmpty(requestAccessEmailValue) && requestAccessEmailValue.Length >= 255)
                        {
                            requestAccessEmailValue = requestAccessEmailValue.Substring(0, 255);
                        }
                        if (!String.IsNullOrEmpty(requestAccessEmailValue))
                        {
                            web.RequestAccessEmail = requestAccessEmailValue;

                            web.Update();
                            web.Context.ExecuteQueryRetry();
                        }
                    }

                    if (!isNoScriptSite)
                    {
                        web.NoCrawl = webSettings.NoCrawl;
                    }
                    else
                    {
                        scope.LogWarning(CoreResources.Provisioning_ObjectHandlers_WebSettings_SkipNoCrawlUpdate);
                    }

                    if (web.CommentsOnSitePagesDisabled != webSettings.CommentsOnSitePagesDisabled)
                    {
                        web.CommentsOnSitePagesDisabled = webSettings.CommentsOnSitePagesDisabled;
                    }

                    if (web.ExcludeFromOfflineClient != webSettings.ExcludeFromOfflineClient)
                    {
                        web.ExcludeFromOfflineClient = webSettings.ExcludeFromOfflineClient;
                    }

                    if (web.MembersCanShare != webSettings.MembersCanShare)
                    {
                        web.MembersCanShare = webSettings.MembersCanShare;
                    }

                    if (web.DisableFlows != webSettings.DisableFlows)
                    {
                        web.DisableFlows = webSettings.DisableFlows;
                    }

                    if (web.DisableAppViews != webSettings.DisableAppViews)
                    {
                        web.DisableAppViews = webSettings.DisableAppViews;
                    }

                    if (web.HorizontalQuickLaunch != webSettings.HorizontalQuickLaunch)
                    {
                        web.HorizontalQuickLaunch = webSettings.HorizontalQuickLaunch;
                    }

                    if (web.SearchScope.ToString() != webSettings.SearchScope.ToString())
                    {
                        web.SearchScope = (SearchScopeType)Enum.Parse(typeof(SearchScopeType), webSettings.SearchScope.ToString(), true);
                    }

                    if (web.SearchBoxInNavBar.ToString() != webSettings.SearchBoxInNavBar.ToString())
                    {
                        web.SearchBoxInNavBar = (SearchBoxInNavBarType)Enum.Parse(typeof(SearchBoxInNavBarType), webSettings.SearchBoxInNavBar.ToString(), true);
                    }

                    var masterUrl = parser.ParseString(webSettings.MasterPageUrl);
                    if (!string.IsNullOrEmpty(masterUrl))
                    {
                        if (!isNoScriptSite)
                        {
                            web.MasterUrl = masterUrl;
                        }
                        else
                        {
                            scope.LogWarning(CoreResources.Provisioning_ObjectHandlers_WebSettings_SkipMasterPageUpdate);
                        }
                    }
                    var customMasterUrl = parser.ParseString(webSettings.CustomMasterPageUrl);
                    if (!string.IsNullOrEmpty(customMasterUrl))
                    {
                        if (!isNoScriptSite)
                        {
                            web.CustomMasterUrl = customMasterUrl;
                        }
                        else
                        {
                            scope.LogWarning(CoreResources.Provisioning_ObjectHandlers_WebSettings_SkipCustomMasterPageUpdate);
                        }
                    }
                    if (!String.IsNullOrEmpty(webSettings.Title))
                    {
                        var newTitle = parser.ParseString(webSettings.Title);
                        if (newTitle != web.Title)
                        {
                            web.Title = newTitle;
                        }
                    }
                    if (!String.IsNullOrEmpty(webSettings.Description))
                    {
                        var newDescription = parser.ParseString(webSettings.Description);
                        if (newDescription != web.Description)
                        {
                            web.Description = newDescription;
                        }
                    }
                    if (webSettings.SiteLogo != null)
                    {
                        var logoUrl = parser.ParseString(webSettings.SiteLogo);
                        if (template.BaseSiteTemplate == "SITEPAGEPUBLISHING#0" && web.WebTemplate == "GROUP")
                        {
                            // logo provisioning throws when applying across base template IDs; provisioning fails in this case
                            // this is the error that is already (rightly so) shown beforehand in the console: WARNING: The source site from which the template was generated had a base template ID value of SITEPAGEPUBLISHING#0, while the current target site has a base template ID value of GROUP#0. This could cause potential issues while applying the template.
                            WriteMessage("Applying site logo across base template IDs is not possible. Skipping site logo provisioning.", ProvisioningMessageType.Warning);
                        }
                        else
                        {
                            // Modern site? Then we assume the SiteLogo is actually a filepath
                            if (web.WebTemplate == "GROUP")
                            {
                                if (!string.IsNullOrEmpty(logoUrl) && !logoUrl.ToLower().Contains("_api/groupservice/getgroupimage"))
                                {
                                    var fileBytes = ConnectorFileHelper.GetFileBytes(template.Connector, logoUrl);
                                    if (fileBytes != null && fileBytes.Length > 0)
                                    {
                                        var mimeType = "";
                                        var imgUrl = logoUrl;
                                        if (imgUrl.Contains("?"))
                                        {
                                            imgUrl = imgUrl.Split(new[] { '?' })[0];
                                        }
                                        if (imgUrl.EndsWith(".gif", StringComparison.InvariantCultureIgnoreCase))
                                        {
                                            mimeType = "image/gif";
                                        }
                                        if (imgUrl.EndsWith(".png", StringComparison.InvariantCultureIgnoreCase))
                                        {
                                            mimeType = "image/png";
                                        }
                                        if (imgUrl.EndsWith(".jpg", StringComparison.InvariantCultureIgnoreCase))
                                        {
                                            mimeType = "image/jpeg";
                                        }
                                        Sites.SiteCollection.SetGroupImageAsync((ClientContext)web.Context, fileBytes, mimeType).GetAwaiter().GetResult();

                                    }
                                }
                            }
                            else
                            {
                                web.SiteLogoUrl = logoUrl;
                            }
                        }
                    }
                    if (webSettings.SiteLogoThumbnail != null)
                    {
                        // SiteLogoThumbnail handling via PnP Core SDK
                        // https://pnp.github.io/pnpcore/using-the-sdk/branding-intro.html#set-and-reset-the-site-logo-and-site-logo-thumbnail

                        // Try to get the logo thumbnail
                        var logoThumbnailUrl = parser.ParseString(webSettings.SiteLogoThumbnail);

                        // And try to get the logo thumbnail filename
                        var logoThumbnailFilename = logoThumbnailUrl.IndexOf('/') > 0 ?
                            logoThumbnailUrl.Substring(logoThumbnailUrl.LastIndexOf('/') + 1) :
                            (logoThumbnailUrl.IndexOf('\\') > 0 ?
                            logoThumbnailUrl.Substring(logoThumbnailUrl.LastIndexOf('\\') + 1) :
                            "logo-thumbnail.jpg"); // __siteIcon__

                        var fileBytes = ConnectorFileHelper.GetFileBytes(template.Connector, logoThumbnailUrl);
                        if (fileBytes != null && fileBytes.Length > 0)
                        {
                            // Get the file stream
                            using (var fileStream = new MemoryStream())
                            {
                                fileStream.Write(fileBytes, 0, fileBytes.Length);
                                fileStream.Position = 0;

                                // Move to the PnP Core SDK context
                                using (var pnpCoreContext = PnPCoreSdk.Instance.GetPnPContext(web.Context as ClientContext))
                                {
                                    // Get the Chrome options
                                    var chrome = pnpCoreContext.Web.GetBrandingManager().GetChromeOptions();

                                    // And configure the new site logo thumbnail
                                    chrome.Header.SetSiteLogoThumbnail(logoThumbnailFilename, fileStream, true);
                                    chrome.Header.HideTitle = !template.Header.ShowSiteTitle || String.IsNullOrEmpty(webSettings.Title);

                                    pnpCoreContext.Web.GetBrandingManager().SetChromeOptions(chrome);
                                }
                            }
                        }
                    }
                    var welcomePage = parser.ParseString(webSettings.WelcomePage);
                    if (!string.IsNullOrEmpty(welcomePage))
                    {
                        if (welcomePage != web.RootFolder.WelcomePage)
                        {
                            web.RootFolder.WelcomePage = welcomePage;
                            web.RootFolder.Update();
                        }
                    }
                    if (!string.IsNullOrEmpty(webSettings.AlternateCSS))
                    {
                        var newAlternateCssUrl = parser.ParseString(webSettings.AlternateCSS);
                        if (newAlternateCssUrl != web.AlternateCssUrl)
                        {
                            web.AlternateCssUrl = newAlternateCssUrl;
                        }
                    }

                    // Temporary disabled as this change is a breaking change for folks that have not set this property in their provisioning templates
                    //web.QuickLaunchEnabled = webSettings.QuickLaunchEnabled;

                    web.Update();
                    web.Context.ExecuteQueryRetry();

                    //GetWebSearchCenterUrl makes us loose the pending changes in CSOM Context
                    string searchCenterUrl = parser.ParseString(webSettings.SearchCenterUrl);
                    if (!string.IsNullOrEmpty(searchCenterUrl) &&
                        web.GetWebSearchCenterUrl(true) != searchCenterUrl)
                    {
                        web.SetWebSearchCenterUrl(searchCenterUrl);
                    }

                    if (webSettings.HubSiteUrl != null)
                    {
                        if (TenantExtensions.IsCurrentUserTenantAdmin(web.Context as ClientContext))
                        {
                            var hubsiteUrl = parser.ParseString(webSettings.HubSiteUrl);
                            try
                            {
                                using (var tenantContext = web.Context.Clone(web.GetTenantAdministrationUrl(), applyingInformation.AccessTokens))
                                {
                                    var tenant = new Tenant(tenantContext);
                                    tenant.ConnectSiteToHubSite(web.Url, hubsiteUrl);
                                    tenantContext.ExecuteQueryRetry();
                                }
                            }
                            catch (Exception ex)
                            {
                                WriteMessage($"Hub site association failed: {ex.Message}", ProvisioningMessageType.Warning);
                            }
                        }
                        else
                        {
                            WriteMessage("You need to be a SharePoint admin when associating to a Hub site.", ProvisioningMessageType.Warning);
                        }
                    }
                }
            }

            return parser;
        }

        public override bool WillExtract(Web web, ProvisioningTemplate template, ProvisioningTemplateCreationInformation creationInfo)
        {
            return true;
        }

        public override bool WillProvision(Web web, ProvisioningTemplate template, ProvisioningTemplateApplyingInformation applyingInformation)
        {
            return template.WebSettings != null;
        }


    }
}
