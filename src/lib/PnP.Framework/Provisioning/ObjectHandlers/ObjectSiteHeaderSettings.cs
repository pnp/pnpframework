using Microsoft.SharePoint.Client;
using PnP.Framework.Diagnostics;
using PnP.Framework.Provisioning.Model;
using System;
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
            using (new PnPMonitoredScope(this.Name))
            {
                web.EnsureProperties(w => w.HeaderEmphasis, w => w.HeaderLayout, w => w.MegaMenuEnabled);

                var header = new SiteHeader
                {
                    MenuStyle = web.MegaMenuEnabled ? SiteHeaderMenuStyle.MegaMenu : SiteHeaderMenuStyle.Cascading
                };

                switch (web.HeaderLayout)
                {
                    case HeaderLayoutType.Compact:
                        {
                            header.Layout = SiteHeaderLayout.Compact;
                            break;
                        }

                    case HeaderLayoutType.Minimal:
                        {
                            header.Layout = SiteHeaderLayout.Minimal;
                            break;
                        }

                    case HeaderLayoutType.Extended:
                        {
                            header.Layout = SiteHeaderLayout.Extended;
                            break;
                        }

                    default:
                        {
                            header.Layout = SiteHeaderLayout.Standard;
                            break;
                        }
                }

                if (Enum.TryParse(web.HeaderEmphasis.ToString(), out Emphasis backgroundEmphasis))
                {
                    header.BackgroundEmphasis = backgroundEmphasis;
                }

                // Move to the PnP Core SDK context
                using (var pnpCoreContext = PnPCoreSdk.Instance.GetPnPContext(web.Context as ClientContext))
                {
                    // Get the Chrome options
                    var chrome = pnpCoreContext.Web.GetBrandingManager().GetChromeOptions();

                    header.ShowSiteTitle = !chrome.Header.HideTitle;
                    header.ShowSiteNavigation = chrome.Navigation.Visible;
                }

                template.Header = header;
            }

            return template;
        }

        public override TokenParser ProvisionObjects(Web web, ProvisioningTemplate template, TokenParser parser, ProvisioningTemplateApplyingInformation applyingInformation)
        {
            using (new PnPMonitoredScope(this.Name))
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
    }
}
