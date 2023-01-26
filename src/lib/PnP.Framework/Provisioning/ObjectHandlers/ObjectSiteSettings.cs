using Microsoft.SharePoint.Client;
using PnP.Framework.Diagnostics;
using PnP.Framework.Provisioning.Model;
using System;

namespace PnP.Framework.Provisioning.ObjectHandlers
{
    internal class ObjectSiteSettings : ObjectHandlerBase
    {
        public override string Name
        {
            get { return "Site Settings"; }
        }

        public override string InternalName => "SiteSettings";
        public override ProvisioningTemplate ExtractObjects(Web web, ProvisioningTemplate template, ProvisioningTemplateCreationInformation creationInfo)
        {
            using (var scope = new PnPMonitoredScope(this.Name))
            {
                // Try to get access to the Site Collection in the current context
                var site = (web.Context as ClientContext)?.Site;
                if (site != null)
                {
                    // And if we have it, load the properties in which we're interested into
                    site.EnsureProperties(
                        s => s.AllowDesigner,
                        s => s.AllowCreateDeclarativeWorkflow,
                        s => s.AllowSaveDeclarativeWorkflowAsTemplate,
                        s => s.AllowSavePublishDeclarativeWorkflow,
                        s => s.SocialBarOnSitePagesDisabled,
                        s => s.SearchBoxInNavBar,
                        s => s.ShowPeoplePickerSuggestionsForGuestUsers
                        );

                    // Configure the output SiteSettings object
                    var siteSettings = new SiteSettings
                    {
                        AllowDesigner = site.AllowDesigner,
                        AllowCreateDeclarativeWorkflow = site.AllowCreateDeclarativeWorkflow,
                        AllowSaveDeclarativeWorkflowAsTemplate = site.AllowSaveDeclarativeWorkflowAsTemplate,
                        AllowSavePublishDeclarativeWorkflow = site.AllowSavePublishDeclarativeWorkflow,
                        SocialBarOnSitePagesDisabled = site.SocialBarOnSitePagesDisabled,
                        SearchBoxInNavBar = (SearchBoxInNavBar)Enum.Parse(typeof(SearchBoxInNavBar), site.SearchBoxInNavBar.ToString()),
                        SearchCenterUrl = site.RootWeb.GetSiteCollectionSearchCenterUrl(),
                        ShowPeoplePickerSuggestionsForGuestUsers = site.ShowPeoplePickerSuggestionsForGuestUsers
                    };

                    // Update the provisioning template accordingly
                    template.SiteSettings = siteSettings;
                }
            }
            return template;
        }

        public override TokenParser ProvisionObjects(Web web, ProvisioningTemplate template, TokenParser parser, ProvisioningTemplateApplyingInformation applyingInformation)
        {
            using (var scope = new PnPMonitoredScope(this.Name))
            {
                if (template.SiteSettings != null)
                {
                    // Try to get access to the Site Collection in the current context
                    var site = (web.Context as ClientContext)?.Site;
                    if (site != null)
                    {
                        // Apply the following properties if and only if the target site is a classic one
                        if (!(site.IsCommunicationSite() || site.IsModernTeamSite()))
                        {
                            site.AllowDesigner = template.SiteSettings.AllowDesigner;
                            site.AllowCreateDeclarativeWorkflow = template.SiteSettings.AllowCreateDeclarativeWorkflow;
                            site.AllowSaveDeclarativeWorkflowAsTemplate = template.SiteSettings.AllowSaveDeclarativeWorkflowAsTemplate;
                            site.AllowSavePublishDeclarativeWorkflow = template.SiteSettings.AllowSavePublishDeclarativeWorkflow;
                            // Call ExecuteQuery right away to ensure it's called with the correct Context as later site.Is*() calls might still trigger ExecuteQuery but on at wrong context
                            web.Context.ExecuteQueryRetryAsync();
                        }

                        // Right now this works in Communication Sites only
                        // For further details: https://github.com/SharePoint/sp-dev-docs/issues/1532
                        if (site.IsCommunicationSite())
                        {
                            site.SocialBarOnSitePagesDisabled = template.SiteSettings.SocialBarOnSitePagesDisabled;
                            // Call ExecuteQuery right away to ensure it's called with the correct Context as later site.Is*() calls might still trigger ExecuteQuery but on at wrong context
                            web.Context.ExecuteQueryRetryAsync();
                        }

                        site.EnsureProperty(s => s.SearchBoxInNavBar);
                        if (site.SearchBoxInNavBar.ToString() != template.SiteSettings.SearchBoxInNavBar.ToString())
                        {
                            site.SearchBoxInNavBar = (SearchBoxInNavBarType)Enum.Parse(typeof(SearchBoxInNavBarType), template.SiteSettings.SearchBoxInNavBar.ToString(), true);
                            // Call ExecuteQuery right away to ensure it's called with the correct Context as later site.Is*() calls might still trigger ExecuteQuery but on at wrong context
                            web.Context.ExecuteQueryRetryAsync();
                        }

                        if (!string.IsNullOrEmpty(template.SiteSettings.SearchCenterUrl) &&
                            site.RootWeb.GetSiteCollectionSearchCenterUrl() != template.SiteSettings.SearchCenterUrl)
                        {
                            site.RootWeb.SetSiteCollectionSearchCenterUrl(template.SiteSettings.SearchCenterUrl);
                        }

                        site.ShowPeoplePickerSuggestionsForGuestUsers = template.SiteSettings.ShowPeoplePickerSuggestionsForGuestUsers;
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
            return template.SiteSettings != null;
        }


    }
}
