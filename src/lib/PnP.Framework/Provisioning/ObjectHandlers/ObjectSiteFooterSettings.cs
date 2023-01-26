﻿using Microsoft.SharePoint.Client;
using Newtonsoft.Json;
using PnP.Framework.Diagnostics;
using PnP.Framework.Provisioning.Model;
using PnP.Framework.Provisioning.ObjectHandlers.Extensions;
using PnP.Framework.Utilities;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Web;

namespace PnP.Framework.Provisioning.ObjectHandlers
{
    internal class ObjectSiteFooterSettings : ObjectHandlerBase
    {
        //const string footerNodeKey = "13b7c916-4fea-4bb2-8994-5cf274aeb530";
        //const string titleNodeKey = "7376cd83-67ac-4753-b156-6a7b3fa0fc1f";
        //const string logoNodeKey = "2e456c2e-3ded-4a6c-a9ea-f7ac4c1b5100";
        //const string menuNodeKey = "3a94b35f-030b-468e-80e3-b75ee84ae0ad";
        public override string Name
        {
            get { return "Site Footer"; }
        }

        public override string InternalName => "SiteFooter";

        public override ProvisioningTemplate ExtractObjects(Web web, ProvisioningTemplate template, ProvisioningTemplateCreationInformation creationInfo)
        {
            using (var scope = new PnPMonitoredScope(this.Name))
            {
                web.EnsureProperties(w => w.FooterEnabled, w => w.ServerRelativeUrl, w => w.Url, w => w.Language);
                var defaultCulture = new CultureInfo((int)web.Language);

                var footer = new SiteFooter
                {
                    Enabled = web.FooterEnabled
                };

                // Move to the PnP Core SDK context
                using (var pnpCoreContext = PnPCoreSdk.Instance.GetPnPContext(web.Context as ClientContext))
                {
                    // Get the Chrome options
                    var chrome = pnpCoreContext.Web.GetBrandingManager().GetChromeOptions();

                    footer.Enabled = chrome.Footer.Enabled;
                    footer.DisplayName = chrome.Footer.DisplayName;
                    footer.Layout = (PnP.Framework.Provisioning.Model.SiteFooterLayout)Enum.Parse(typeof(PnP.Framework.Provisioning.Model.SiteFooterLayout), chrome.Footer.Layout.ToString());
                    footer.BackgroundEmphasis = (PnP.Framework.Provisioning.Model.Emphasis)Enum.Parse(typeof(PnP.Framework.Provisioning.Model.Emphasis), chrome.Footer.Emphasis.ToString());
                }

                var structureString = web.ExecuteGetAsync($"/_api/navigation/MenuState?menuNodeKey='{Constants.SITEFOOTER_NODEKEY}'", defaultCulture.Name).GetAwaiter().GetResult();
                var menuState = JsonConvert.DeserializeObject<MenuState>(structureString);

                if (menuState.Nodes.Count > 0)
                {
                    // Find the title node
                    var titleNode = menuState.Nodes.FirstOrDefault(n => n.Title == Constants.SITEFOOTER_TITLENODEKEY);
                    if (titleNode != null)
                    {
                        var titleNodeNodes = titleNode.Nodes;
                        if (titleNodeNodes.Count > 0)
                        {
                            if (!string.IsNullOrEmpty(titleNodeNodes[0].SimpleUrl))
                            {
                                footer.Logo = Tokenize(titleNodeNodes[0].SimpleUrl, web.ServerRelativeUrl);
                            }
                            if (!string.IsNullOrEmpty(titleNodeNodes[0].Title))
                            {
                                if (creationInfo.PersistMultiLanguageResources)
                                {
                                    if (UserResourceExtensions.PersistResourceValue($"FooterNavigationNode_{titleNode.Key}_{titleNodeNodes[0].Key}_Title", defaultCulture.LCID, titleNodeNodes[0].Title, creationInfo))
                                    {
                                        footer.Name = $"{{res:FooterNavigationNode_{titleNode.Key}_{titleNodeNodes[0].Key}_Title}}";
                                    }
                                }
                                else
                                {
                                    footer.Name = titleNodeNodes[0].Title;
                                }
                            }
                        }
                    }
                    // find the logo node
                    if (string.IsNullOrEmpty(footer.Logo))
                    {
                        var logoNode = menuState.Nodes.FirstOrDefault(n => n.Title == Constants.SITEFOOTER_LOGONODEKEY);
                        if (logoNode != null)
                        {
                            footer.Logo = Tokenize(logoNode.SimpleUrl, web.ServerRelativeUrl);
                        }
                    }
                }
                // find the menu Nodes
                var menuNodesNode = menuState.Nodes.FirstOrDefault(n => n.Title == Constants.SITEFOOTER_MENUNODEKEY);
                if (menuNodesNode != null)
                {
                    foreach (var innerMenuNode in menuNodesNode.Nodes)
                    {
                        footer.FooterLinks.Add(ParseNodes(innerMenuNode, template, web.ServerRelativeUrl, creationInfo.PersistMultiLanguageResources, defaultCulture, menuNodesNode.Key, creationInfo));
                    }
                }
                if (creationInfo.ExtractConfiguration != null && creationInfo.ExtractConfiguration.SiteFooter != null && creationInfo.ExtractConfiguration.SiteFooter.RemoveExistingNodes)
                {
                    footer.RemoveExistingNodes = true;
                }

                if (creationInfo.PersistMultiLanguageResources)
                {
                    //get Titles for the rest of the Languages
                    foreach (var language in template.SupportedUILanguages.Where(c => c.LCID != defaultCulture.LCID))
                    {
                        var currentCulture = new CultureInfo(language.LCID);
                        var structureStringMUI = web.ExecuteGetAsync($"/_api/navigation/MenuState?menuNodeKey='{Constants.SITEFOOTER_NODEKEY}'", currentCulture.Name).GetAwaiter().GetResult();
                        var menuStateMUI = JsonConvert.DeserializeObject<MenuState>(structureStringMUI);

                        if (menuStateMUI.Nodes.Count > 0)
                        {
                            var titleNode = menuStateMUI.Nodes.FirstOrDefault(n => n.Title == Constants.SITEFOOTER_TITLENODEKEY);
                            if (titleNode != null)
                            {
                                var titleNodeNodes = titleNode.Nodes;
                                if (titleNodeNodes.Count > 0)
                                {
                                    if (!string.IsNullOrEmpty(titleNodeNodes[0].Title))
                                    {
                                        if (UserResourceExtensions.PersistResourceValue($"FooterNavigationNode_{titleNode.Key}_{titleNodeNodes[0].Key}_Title", currentCulture.LCID, titleNodeNodes[0].Title, creationInfo))
                                        {
                                            footer.Name = $"{{res:FooterNavigationNode_{titleNode.Key}_{titleNodeNodes[0].Key}_Title}}";
                                        }
                                    }
                                }
                            }
                        }
                        // find the menu Nodes

                        var menuNodesNodeMUI = menuStateMUI.Nodes.FirstOrDefault(n => n.Title == Constants.SITEFOOTER_MENUNODEKEY);
                        if (menuNodesNodeMUI != null)
                        {
                            foreach (var innerMenuNode in menuNodesNodeMUI.Nodes)
                            {
                                ParseNodesMUI(innerMenuNode, web.ServerRelativeUrl, currentCulture, menuNodesNode.Key, creationInfo);
                            }
                        }
                    }
                }

                template.Footer = footer;
                if (creationInfo.PersistBrandingFiles)
                {
                    // Extract site logo if property has been set and it's not dynamic image from _api URL
                    if (!string.IsNullOrEmpty(template.Footer.Logo) && (!template.Footer.Logo.ToLowerInvariant().Contains("_api/")))
                    {
                        // Convert to server relative URL if needed (can be set to FQDN URL of a file hosted in the site (e.g. for communication sites))
                        Uri webUri = new Uri(web.Url);
                        string webUrl = $"{webUri.Scheme}://{webUri.DnsSafeHost}";
                        string footerLogoServerRelativeUrl = template.Footer.Logo.Replace(webUrl, "");

                        if (PersistFile(web, creationInfo, scope, footerLogoServerRelativeUrl))
                        {
                            template.Files.Add(GetTemplateFile(web, footerLogoServerRelativeUrl));
                        }
                    }
                    template.Footer.Logo = Tokenize(template.Footer.Logo, web.Url, web);
                    var files = template.Files.Distinct().ToList();
                    template.Files.Clear();
                    template.Files.AddRange(files);
                }
            }
            return template;
        }

        private void ParseNodesMUI(MenuNode node, string webServerRelativeUrl, CultureInfo currentCulture, string parentKey, ProvisioningTemplateCreationInformation creationInfo)
        {
            UserResourceExtensions.PersistResourceValue($"FooterNavigationNode_{parentKey}_{node.Key}_Title", currentCulture.LCID, node.Title, creationInfo);

            if (node.Nodes.Count > 0)
            {
                foreach (var childNode in node.Nodes)
                {
                    ParseNodesMUI(childNode, webServerRelativeUrl, currentCulture, node.Key, creationInfo);
                }
            }
        }
        private string TokenizeHost(Web web, string json)
        {
            // HostUrl token replacement
            var uri = new Uri(web.Url);
            json = Regex.Replace(json, $"{uri.Scheme}://{uri.DnsSafeHost}:{uri.Port}", "{hosturl}", RegexOptions.IgnoreCase);
            json = Regex.Replace(json, $"{uri.Scheme}://{uri.DnsSafeHost}", "{hosturl}", RegexOptions.IgnoreCase);
            return json;
        }

        //TODO: Candidate for cleanup
        private Model.File GetTemplateFile(Web web, string serverRelativeUrl)
        {

            var webServerUrl = web.EnsureProperty(w => w.Url);
            var serverUri = new Uri(webServerUrl);
            var serverUrl = $"{serverUri.Scheme}://{serverUri.Authority}";
            var fullUri = new Uri(UrlUtility.Combine(serverUrl, serverRelativeUrl));

            var folderPath = fullUri.Segments.Take(fullUri.Segments.Length - 1).ToArray().Aggregate((i, x) => i + x).TrimEnd('/');
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
                        web.Context.Load(file);
                        web.Context.ExecuteQueryRetry();
                        ClientResult<Stream> stream = file.OpenBinaryStream();
                        web.Context.ExecuteQueryRetry();
                        
                        file.EnsureProperty(f => f.ServerRelativePath);
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
                    WriteMessage("No connector present to persist footer logo.", ProvisioningMessageType.Error);
                    scope.LogError("No connector present to persist footer logo.");
                }
            }
            else
            {
                success = true;
            }
            return success;
        }

        private void CopyStream(Stream source, Stream destination)
        {
            byte[] buffer = new byte[32768];
            int bytesRead;

            do
            {
                bytesRead = source.Read(buffer, 0, buffer.Length);
                destination.Write(buffer, 0, bytesRead);
            } while (bytesRead != 0);
        }
        private SiteFooterLink ParseNodes(MenuNode node, ProvisioningTemplate template, string webServerRelativeUrl, bool persistLanguage, CultureInfo currentCulture, string parentKey, ProvisioningTemplateCreationInformation creationInfo)
        {
            var link = new SiteFooterLink();

            if (persistLanguage)
            {
                if (UserResourceExtensions.PersistResourceValue($"FooterNavigationNode_{parentKey}_{node.Key}_Title", currentCulture.LCID, node.Title, creationInfo))
                {
                    link.DisplayName = $"{{res:FooterNavigationNode_{parentKey}_{node.Key}_Title}}";
                }
            }
            else
            {
                link.DisplayName = node.Title;
            }

            link.Url = Tokenize(node.SimpleUrl, webServerRelativeUrl);

            if (node.Nodes.Count > 0)
            {
                link.FooterLinks = new SiteFooterLinkCollection(template);
                foreach (var childNode in node.Nodes)
                {
                    link.FooterLinks.Add(ParseNodes(childNode, template, webServerRelativeUrl, persistLanguage, currentCulture, node.Key, creationInfo));
                }
            }
            return link;
        }

        public override TokenParser ProvisionObjects(Web web, ProvisioningTemplate template, TokenParser parser, ProvisioningTemplateApplyingInformation applyingInformation)
        {
            using (var scope = new PnPMonitoredScope(this.Name))
            {
                if (template.Footer != null)
                {
                    web.EnsureProperties(w => w.ServerRelativeUrl,
                        w => w.FooterEnabled,
                        w => w.FooterLayout,
                        w => w.FooterEmphasis,
                        w => w.Language);
                    web.FooterEnabled = template.Footer.Enabled;
                    var defaultCulture = new CultureInfo((int)web.Language);

                    //var jsonRequest = new
                    //{
                    //    footerEnabled = web.FooterEnabled,
                    //    footerLayout = web.FooterLayout,
                    //    footerEmphasis = web.FooterEmphasis
                    //};

                    //web.ExecutePostAsync("/_api/web/SetChromeOptions", System.Text.Json.JsonSerializer.Serialize(jsonRequest)).GetAwaiter().GetResult();

                    // Move to the PnP Core SDK context
                    using (var pnpCoreContext = PnPCoreSdk.Instance.GetPnPContext(web.Context as ClientContext))
                    {
                        // Get the Chrome options
                        var chrome = pnpCoreContext.Web.GetBrandingManager().GetChromeOptions();

                        chrome.Footer.Enabled = web.FooterEnabled;
                        chrome.Footer.DisplayName = template.Footer.DisplayName;
                        chrome.Footer.Layout = (PnP.Core.Model.SharePoint.FooterLayoutType)Enum.Parse(typeof(PnP.Core.Model.SharePoint.FooterLayoutType), template.Footer.Layout.ToString());
                        chrome.Footer.Emphasis = (PnP.Core.Model.SharePoint.FooterVariantThemeType)Enum.Parse(typeof(PnP.Core.Model.SharePoint.FooterVariantThemeType), template.Footer.BackgroundEmphasis.ToString()); 

                        pnpCoreContext.Web.GetBrandingManager().SetChromeOptions(chrome);
                    }

                    //if (PnPProvisioningContext.Current != null)
                    //{
                    //    // Get an Access Token for the SetChromeOptions request
                    //    var spoResourceUri = new Uri(web.Url).Authority;
                    //    var accessToken = PnPProvisioningContext.Current.AcquireToken(spoResourceUri, null);

                    //    if (accessToken != null)
                    //    {
                    //        // Prepare the JSON request for SetChromeOptions
                    //        var jsonRequest = new
                    //        {
                    //            footerEnabled = web.FooterEnabled,
                    //            footerLayout = web.FooterLayout,
                    //            footerEmphasis = web.FooterEmphasis
                    //        };

                    //        // Build the URL of the SetChromeOptions API
                    //        var setChromeOptionsApiUrl = $"{web.Url}/_api/web/SetChromeOptions";

                    //        // Make the POST request to the SetChromeOptions API
                    //        // and fail in case of any exception
                    //        HttpHelper.MakePostRequest(setChromeOptionsApiUrl,
                    //            jsonRequest,
                    //            "application/json",
                    //            accessToken);
                    //    }
                    //}
                    //else
                    //{
                    //    web.Update();
                    //    web.Context.ExecuteQueryRetry();
                    //}

                    if (web.FooterEnabled)
                    {
                        var structureString = web.ExecuteGetAsync($"/_api/navigation/MenuState?menuNodeKey='{Constants.SITEFOOTER_NODEKEY}'", defaultCulture.Name).GetAwaiter().GetResult();
                        var menuState = JsonConvert.DeserializeObject<MenuState>(structureString);
                        if (menuState.StartingNodeKey == null)
                        {

                            var now = DateTime.Now.ToString("yyyy-MM-ddTHH:mm:ss:Z");
                            web.ExecutePostAsync($"/_api/navigation/SaveMenuState", $@"{{ ""menuState"":{{ ""Version"":""{now}"",""StartingNodeTitle"":""3a94b35f-030b-468e-80e3-b75ee84ae0ad"",""SPSitePrefix"":""/"",""SPWebPrefix"":""{web.ServerRelativeUrl}"",""FriendlyUrlPrefix"":"""",""SimpleUrl"":"""",""Nodes"":[]}}}}", defaultCulture.Name).GetAwaiter().GetResult();
                            structureString = web.ExecuteGetAsync($"/_api/navigation/MenuState?menuNodeKey='{Constants.SITEFOOTER_NODEKEY}'", defaultCulture.Name).GetAwaiter().GetResult();
                            menuState = JsonConvert.DeserializeObject<MenuState>(structureString);
                        }
                        var n1 = web.Navigation.GetNodeById(Convert.ToInt32(menuState.StartingNodeKey));

                        web.Context.Load(n1);
                        web.Context.ExecuteQueryRetry();

                        web.Context.Load(n1, n => n.Children.IncludeWithDefaultProperties());
                        web.Context.ExecuteQueryRetry();

                        var menuNode = n1.Children.FirstOrDefault(n => n.Title == Constants.SITEFOOTER_MENUNODEKEY);
                        if (menuNode != null)
                        {
                            if (template.Footer.RemoveExistingNodes == true)
                            {
                                menuNode.DeleteObject();
                                web.Context.ExecuteQueryRetry();

                                menuNode = n1.Children.Add(new NavigationNodeCreationInformation()
                                {
                                    Title = Constants.SITEFOOTER_MENUNODEKEY
                                });
                            }
                        }
                        else
                        {
                            menuNode = n1.Children.Add(new NavigationNodeCreationInformation()
                            {
                                Title = Constants.SITEFOOTER_MENUNODEKEY
                            });
                        }

                        if (template.Footer.FooterLinks != null && template.Footer.FooterLinks.Any())
                        {
                            for (var q = template.Footer.FooterLinks.Count - 1; q >= 0; q--)
                            {
                                var footerLink = template.Footer.FooterLinks[q];
                                var newParentNode = menuNode.Children.Add(new NavigationNodeCreationInformation()
                                {
                                    Url = ObjectNavigation.ReplaceFileUniqueToken(web, parser.ParseString(footerLink.Url)),
                                    Title = parser.ParseString(footerLink.DisplayName)
                                });

                                if (footerLink.DisplayName.ContainsResourceToken())
                                {
                                    web.Context.ExecuteQueryRetry();
                                    newParentNode.LocalizeNavigationNode(web, footerLink.DisplayName, parser, scope);
                                }
                                if (footerLink.FooterLinks != null && footerLink.FooterLinks.Any())
                                {
                                    for (var s = footerLink.FooterLinks.Count - 1; s >= 0; s--)
                                    {
                                        var childFooterLink = footerLink.FooterLinks[s];
                                        var newChildNode = newParentNode.Children.Add(new NavigationNodeCreationInformation()
                                        {
                                            Url = parser.ParseString(childFooterLink.Url),
                                            Title = parser.ParseString(childFooterLink.DisplayName)
                                        });

                                        if (childFooterLink.DisplayName.ContainsResourceToken())
                                        {
                                            web.Context.ExecuteQueryRetry();
                                            newChildNode.LocalizeNavigationNode(web, childFooterLink.DisplayName, parser, scope);
                                        }
                                    }                                    
                                }
                            }
                        }

                        if (web.Context.PendingRequestCount() > 0)
                        {
                            web.Context.ExecuteQueryRetry();
                        }

                        var logoNode = n1.Children.FirstOrDefault(n => n.Title == Constants.SITEFOOTER_LOGONODEKEY);
                        if (logoNode != null)
                        {
                            if (string.IsNullOrEmpty(template.Footer.Logo))
                            {
                                // remove the logo
                                logoNode.DeleteObject();
                            }
                            else
                            {
                                logoNode.Url = parser.ParseString(template.Footer.Logo);
                            }
                        }
                        else
                        {
                            if (!string.IsNullOrEmpty(template.Footer.Logo))
                            {
                                logoNode = n1.Children.Add(new NavigationNodeCreationInformation()
                                {
                                    Title = Constants.SITEFOOTER_LOGONODEKEY,
                                    Url = parser.ParseString(template.Footer.Logo)
                                });
                            }
                        }
                        if (web.Context.PendingRequestCount() > 0)
                        {
                            web.Context.ExecuteQueryRetry();
                        }

                        var titleNode = n1.Children.FirstOrDefault(n => n.Title == Constants.SITEFOOTER_TITLENODEKEY);
                        if (titleNode != null)
                        {
                            titleNode.EnsureProperty(n => n.Children);
                            if (string.IsNullOrEmpty(template.Footer.Name))
                            {
                                // remove the title
                                titleNode.DeleteObject();
                            }
                            else
                            {
                                titleNode.Children[0].Title = parser.ParseString(template.Footer.Name);
                                titleNode.Update();
                                if (template.Footer.Name.ContainsResourceToken())
                                {
                                    web.Context.ExecuteQueryRetry();
                                    titleNode.LocalizeNavigationNode(web, template.Footer.Name, parser, scope);
                                }
                            }
                        }
                        else
                        {
                            if (!string.IsNullOrEmpty(template.Footer.Name))
                            {
                                titleNode = n1.Children.Add(new NavigationNodeCreationInformation() { Title = Constants.SITEFOOTER_TITLENODEKEY });
                                var node = titleNode.Children.Add(new NavigationNodeCreationInformation() { Title = parser.ParseString(template.Footer.Name) });
                                if (template.Footer.Name.ContainsResourceToken())
                                {
                                    web.Context.ExecuteQueryRetry();
                                    node.LocalizeNavigationNode(web, template.Footer.Name, parser, scope);
                                }
                            }
                        }
                        if (web.Context.PendingRequestCount() > 0)
                        {
                            web.Context.ExecuteQueryRetry();
                        }
                    }
                }
            }
            return parser;
        }

        public override bool WillExtract(Web web, ProvisioningTemplate template, ProvisioningTemplateCreationInformation creationInfo)
        {
            if ((web.Context as ClientContext).Site.IsCommunicationSite())
            {
                return true;
            }
            else
            {
                return false;
            }
        }

        public override bool WillProvision(Web web, ProvisioningTemplate template, ProvisioningTemplateApplyingInformation applyingInformation)
        {
            if ((web.Context as ClientContext).Site.IsCommunicationSite())
            {
                return template.Footer != null;
            }
            else
            {
                return false;
            }
        }

        private class MenuState
        {
            public string FriendlyUrlPrefix { get; set; }
            public List<MenuNode> Nodes { get; set; }

            public string SimpleUrl { get; set; }
            public string SPSitePrefix { get; set; }
            public string SPWebPrefix { get; set; }
            public string StartingNodeKey { get; set; }
            public string StartingNodeTitle { get; set; }
            public string Version { get; set; }

            public MenuState()
            {
                Nodes = new List<MenuNode>();
            }
        }

        private class MenuNode
        {
            public string FriendlyUrlSegment { get; set; }
            public bool IsDeleted { get; set; }
            public bool IsHidden { get; set; }
            public string Key { get; set; }
            public List<MenuNode> Nodes { get; set; }
            public int NodeType { get; set; }
            public string SimpleUrl { get; set; }
            public string Title { get; set; }

            public MenuNode()
            {
                Nodes = new List<MenuNode>();
            }
        }

        private class MenuStateWrapper
        {
            [JsonProperty("menuState")]
            public MenuState MenuState { get; set; }
        }
    }
}
