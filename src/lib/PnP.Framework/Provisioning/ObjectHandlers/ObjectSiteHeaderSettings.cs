using Microsoft.SharePoint.Client;
using PnP.Core.Services;
using PnP.Framework.Diagnostics;
using PnP.Framework.Provisioning.Model;
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
        private PnPContext pnpContext;

        public override ProvisioningTemplate ExtractObjects(Web web, ProvisioningTemplate template, ProvisioningTemplateCreationInformation creationInfo)
        {
            using (new PnPMonitoredScope(this.Name))
            {
                pnpContext = PnPCoreSdk.Instance.GetPnPContext(web.Context as ClientContext);

                pnpContext.Web.EnsureProperties(w => w.HeaderEmphasis, w => w.HeaderLayout, w => w.HideTitleInHeader, w => w.MegaMenuEnabled);

                var header = new SiteHeader
                {
                    MenuStyle = pnpContext.Web.MegaMenuEnabled ? SiteHeaderMenuStyle.MegaMenu : SiteHeaderMenuStyle.Cascading,
                    ShowSiteTitle = !pnpContext.Web.HideTitleInHeader
                };
                switch (pnpContext.Web.HeaderLayout)
                {
                    case Core.Model.SharePoint.HeaderLayoutType.Compact:
                        {
                            header.Layout = SiteHeaderLayout.Compact;
                            break;
                        }
                    case Core.Model.SharePoint.HeaderLayoutType.Minimal:
                        {
                            header.Layout = SiteHeaderLayout.Minimal;
                            break;
                        }
                    case Core.Model.SharePoint.HeaderLayoutType.Extended:
                        {
                            header.Layout = SiteHeaderLayout.Extended;
                            break;
                        }
                    case Core.Model.SharePoint.HeaderLayoutType.Standard:
                    default:
                        {
                            header.Layout = SiteHeaderLayout.Standard;
                            break;
                        }
                }

                if (Enum.TryParse<Emphasis>(pnpContext.Web.HeaderEmphasis.ToString(), out Emphasis backgroundEmphasis))
                {
                    header.BackgroundEmphasis = backgroundEmphasis;
                }

                template.Header = header;
            }
            return template;
        }

        public override TokenParser ProvisionObjects(Web web, ProvisioningTemplate template, TokenParser parser, ProvisioningTemplateApplyingInformation applyingInformation)
        {
            using (var scope = new PnPMonitoredScope(this.Name))
            {
                if (template.Header != null)
                {
                    pnpContext = PnPCoreSdk.Instance.GetPnPContext(web.Context as ClientContext);
                    pnpContext.Web.EnsureProperties(w => w.Url);

                    switch (template.Header.Layout)
                    {
                        case SiteHeaderLayout.Compact:
                            {
                                pnpContext.Web.HeaderLayout = Core.Model.SharePoint.HeaderLayoutType.Compact;
                                break;
                            }
                        case SiteHeaderLayout.Minimal:
                            {
                                pnpContext.Web.HeaderLayout = Core.Model.SharePoint.HeaderLayoutType.Minimal;
                                break;
                            }
                        case SiteHeaderLayout.Extended:
                            {
                                pnpContext.Web.HeaderLayout = Core.Model.SharePoint.HeaderLayoutType.Extended;
                                break;
                            }
                        case SiteHeaderLayout.Standard:
                        default:
                            {
                                pnpContext.Web.HeaderLayout = Core.Model.SharePoint.HeaderLayoutType.Standard;
                                break;
                            }
                    }
                    pnpContext.Web.HeaderEmphasis = (Core.Model.SharePoint.VariantThemeType)Enum.Parse(typeof(Core.Model.SharePoint.VariantThemeType), template.Header.BackgroundEmphasis.ToString());
                    pnpContext.Web.MegaMenuEnabled = template.Header.MenuStyle == SiteHeaderMenuStyle.MegaMenu;
                    pnpContext.Web.HideTitleInHeader = !template.Header.ShowSiteTitle;
                    pnpContext.Web.Update();
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
