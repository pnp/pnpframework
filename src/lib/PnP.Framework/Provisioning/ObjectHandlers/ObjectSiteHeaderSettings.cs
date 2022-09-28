using Microsoft.SharePoint.Client;
using PnP.Framework.Diagnostics;
using PnP.Framework.Provisioning.Model;
using PnP.Framework.Utilities;
using System;

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

                web.EnsureProperties(w => w.Url, w => w.HeaderLayout);

                switch (template.Header.Layout)
                {
                    case SiteHeaderLayout.Compact:
                        {
                            web.HeaderLayout = HeaderLayoutType.Compact;
                            break;
                        }

                    case SiteHeaderLayout.Minimal:
                        {
                            web.HeaderLayout = HeaderLayoutType.Minimal;
                            break;
                        }

                    case SiteHeaderLayout.Extended:
                        {
                            web.HeaderLayout = HeaderLayoutType.Extended;
                            break;
                        }

                    case SiteHeaderLayout.Standard:
                        {
                            web.HeaderLayout = HeaderLayoutType.Standard;
                            break;
                        }
                }

                web.HeaderEmphasis = (SPVariantThemeType)Enum.Parse(typeof(SPVariantThemeType), template.Header.BackgroundEmphasis.ToString());
                web.MegaMenuEnabled = template.Header.MenuStyle == SiteHeaderMenuStyle.MegaMenu;
                web.HideTitleInHeader = !template.Header.ShowSiteTitle;

                var jsonRequest = new
                {
                    headerLayout = web.HeaderLayout,
                    headerEmphasis = web.HeaderEmphasis,
                    megaMenuEnabled = web.MegaMenuEnabled,
                    hideTitleInHeader = web.HideTitleInHeader
                };
				
                web.ExecutePostAsync("/_api/web/SetChromeOptions", System.Text.Json.JsonSerializer.Serialize(jsonRequest)).GetAwaiter().GetResult();

                // Move to the PnP Core SDK context
                using (var pnpCoreContext = PnPCoreSdk.Instance.GetPnPContext(web.Context as ClientContext))
                {
                    // Get the Chrome options
                    var chrome = pnpCoreContext.Web.GetBrandingManager().GetChromeOptions();

                    chrome.Header.HideTitle = !template.Header.ShowSiteTitle;
                    chrome.Navigation.Visible = template.Header.ShowSiteNavigation;

                    pnpCoreContext.Web.GetBrandingManager().SetChromeOptions(chrome);
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
