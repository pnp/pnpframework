using Microsoft.SharePoint.Client;
using PnP.Framework.Diagnostics;
using PnP.Framework.Provisioning.Model;
using PnP.Framework.Utilities;
using System;
using System.Linq;
using System.Text.Json;

namespace PnP.Framework.Provisioning.ObjectHandlers
{
    internal class ObjectSiteHeaderSettings : ObjectHandlerBase
    {
        public override string Name
        {
            get { return "Site Header"; }
        }

        public override string InternalName => "SiteHeader";

        public override ProvisioningTemplate ExtractObjects(Web web, ProvisioningTemplate template, ProvisioningTemplateCreationInformation creationInfo)
        {
            using (var scope = new PnPMonitoredScope(this.Name))
            {
                web.EnsureProperties(w => w.AllProperties, w => w.Url);

                // Move to the PnP Core SDK context
                using (var pnpCoreContext = PnPCoreSdk.Instance.GetPnPContext(web.Context as ClientContext))
                {
                    // Get the Chrome options
                    var chrome = pnpCoreContext.Web.GetBrandingManager().GetChromeOptions();
                    var header = new SiteHeader
                    {
                        ShowSiteTitle = !chrome.Header.HideTitle,
                        ShowSiteNavigation = chrome.Navigation.Visible,
                        MenuStyle = chrome.Navigation.MegaMenuEnabled ? SiteHeaderMenuStyle.MegaMenu : SiteHeaderMenuStyle.Cascading
                    };

                    switch(chrome.Header.Layout)
                                            {
                        case Core.Model.SharePoint.HeaderLayoutType.Compact:
                            header.Layout = SiteHeaderLayout.Compact;
                            break;
                        case Core.Model.SharePoint.HeaderLayoutType.Minimal:
                            header.Layout = SiteHeaderLayout.Minimal;
                            break;
                        case Core.Model.SharePoint.HeaderLayoutType.Extended:
                            header.Layout = SiteHeaderLayout.Extended;
                            break;
                        default:
                            header.Layout = SiteHeaderLayout.Standard;
                            break;
                    }

                    switch(chrome.Header.Emphasis)
                    {
                        case Core.Model.SharePoint.VariantThemeType.Neutral:
                            header.BackgroundEmphasis = Emphasis.Neutral;
                            break;
                        case Core.Model.SharePoint.VariantThemeType.Soft:
                            header.BackgroundEmphasis = Emphasis.Soft;
                            break;
                        case Core.Model.SharePoint.VariantThemeType.Strong:
                            header.BackgroundEmphasis = Emphasis.Strong;
                            break;
                        default:
                            header.BackgroundEmphasis = Emphasis.None;
                            break;
                    }
                    template.Header = header;
                }

                if (creationInfo.PersistBrandingFiles)
                {
                    //Header Background Image
                    var backgroundImageUrl = web.GetPropertyBagValueString("BackgroundImageUrl", "");
                    if (!string.IsNullOrWhiteSpace(backgroundImageUrl))
                    {
                        Uri webUri = new Uri(web.Url);
                        string webUrl = $"{webUri.Scheme}://{webUri.DnsSafeHost}";
                        backgroundImageUrl = backgroundImageUrl.Replace(webUrl, "");

                        if (Utilities.FileUtilities.PersistFile(web, creationInfo, scope, this, backgroundImageUrl))
                        {
                            template.Files.Add(GetTemplateFile(web, backgroundImageUrl));
                        }

                        var files = template.Files.Distinct().ToList();
                        template.Files.Clear();
                        template.Files.AddRange(files);
                    }
                }
            }

            return template;
        }

        public override TokenParser ProvisionObjects(Web web, ProvisioningTemplate template, TokenParser parser, ProvisioningTemplateApplyingInformation applyingInformation)
        {
            using (var scope=new PnPMonitoredScope(this.Name))
            {
                if (template.Header == null)
                {
                    return parser;
                }

                web.EnsureProperties(w => w.Url);

                // Move to the PnP Core SDK context
                using (var pnpCoreContext = PnPCoreSdk.Instance.GetPnPContext(web.Context as ClientContext))
                {
                    // Get the Chrome options
                    var brandingManager = pnpCoreContext.Web.GetBrandingManager();

                    var chrome = brandingManager.GetChromeOptions();

                    chrome.Header.HideTitle = !template.Header.ShowSiteTitle;
                    switch (template.Header.Layout)
                    {
                        case SiteHeaderLayout.Compact:
                            chrome.Header.Layout = Core.Model.SharePoint.HeaderLayoutType.Compact;
                            break;
                        case SiteHeaderLayout.Minimal:
                            chrome.Header.Layout = Core.Model.SharePoint.HeaderLayoutType.Minimal;
                            break;
                        case SiteHeaderLayout.Extended:
                            chrome.Header.Layout = Core.Model.SharePoint.HeaderLayoutType.Extended;
                            break;
                        default:
                            chrome.Header.Layout = Core.Model.SharePoint.HeaderLayoutType.Standard;
                            break;
                    }

                    switch (template.Header.BackgroundEmphasis)
                        {
                        case Emphasis.Neutral:
                            chrome.Header.Emphasis = Core.Model.SharePoint.VariantThemeType.Neutral;
                            break;
                        case Emphasis.Soft:
                            chrome.Header.Emphasis = Core.Model.SharePoint.VariantThemeType.Soft;
                            break;
                        case Emphasis.Strong:
                            chrome.Header.Emphasis = Core.Model.SharePoint.VariantThemeType.Strong;
                            break;
                        default:
                            chrome.Header.Emphasis = Core.Model.SharePoint.VariantThemeType.None;
                            break;
                    }
                    chrome.Navigation.Visible = template.Header.ShowSiteNavigation;
                    chrome.Navigation.MegaMenuEnabled = template.Header.MenuStyle == SiteHeaderMenuStyle.MegaMenu;

                    //Modern header settings
                    if(template?.PropertyBagEntries !=null)
                    {
                        foreach (var entry in template.PropertyBagEntries)
                        {
                            if (string.IsNullOrWhiteSpace(entry.Value))
                            {
                                continue;
                            }
                            switch (entry.Key.ToLower())
                            {
                                case "headeroverlaycolor":
                                    {
                                        if (int.TryParse(entry.Value, out var headerOverlayColor) && Enum.IsDefined(typeof(OverlayColorType), headerOverlayColor))
                                        {
                                            chrome.Header.OverlayColor = (Core.Model.SharePoint.OverlayColorType)(OverlayColorType)headerOverlayColor;
                                        }
                                        break;
                                    }
                                case "headeroverlayopacity":
                                    {
                                        if (int.TryParse(entry.Value, out var headerOverlayOpacity))
                                        {
                                            chrome.Header.OverlayOpacity = headerOverlayOpacity;
                                        }
                                        break;
                                    }
                                case "headeroverlaygradientdirection":
                                    {
                                        if (int.TryParse(entry.Value, out var headerOverlayGradientDirection) && Enum.IsDefined(typeof(Core.Model.SharePoint.OverlayGradientDirectionType), headerOverlayGradientDirection))
                                        {
                                            chrome.Header.OverlayGradientDirection = (Core.Model.SharePoint.OverlayGradientDirectionType)headerOverlayGradientDirection;
                                        }
                                        break;
                                    }
                                case "headercolorindexinlightmode":
                                    {
                                        if (int.TryParse(entry.Value, out var headercolorindexinlightmode))
                                        {
                                            chrome.Header.ColorIndexInLightMode = headercolorindexinlightmode;
                                        }
                                        break;
                                    }
                                case "headercolorindexindarkmode":
                                    {
                                        if (int.TryParse(entry.Value, out var headercolorindexindarkmode))
                                        {
                                            chrome.Header.ColorIndexInDarkMode = headercolorindexindarkmode;
                                        }
                                        break;
                                    }
                                case "fontoptionforsitetitle":
                                    {
                                        chrome.Font.SiteTitle = JsonSerializer.Deserialize<Utilities.FontOption>(entry.Value);
                                        break;
                                    }
                                case "fontoptionforsitenav":
                                    {
                                        chrome.Font.SiteNav = JsonSerializer.Deserialize<Utilities.FontOption>(entry.Value);
                                        break;
                                    }

                            }
                        }
                    }

                    brandingManager.SetChromeOptions(chrome);

                    //modern header background image
                    if (template?.PropertyBagEntries != null)
                    {
                        try
                        {
                            var backgroundImageUrl = template.PropertyBagEntries.FirstOrDefault(p => "backgroundimageurl" == p.Key.ToLower())?.Value;
                            if (!string.IsNullOrWhiteSpace(backgroundImageUrl))
                            {
                                var file = template.Files.FirstOrDefault(f => backgroundImageUrl.EndsWith(f.Src, StringComparison.InvariantCultureIgnoreCase));
                                if (file != null)
                                {
                                    var fileName = System.IO.Path.GetFileName(backgroundImageUrl);
                                    chrome.Header.SetHeaderBackgroundImage(fileName, Utilities.FileUtilities.GetFileStream(template, file), overwrite: true);
                                }
                            }
                        }
                        catch (Exception ex)
                        {
                            // Swallowing the exception as this is not a critical operation
                            // and we don't want to fail the whole provisioning
                            scope.LogWarning($"An error occurred while setting the header background image: {ex.Message}");
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
            return template.Header != null;
        }

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
    }
}
