﻿using Microsoft.Online.SharePoint.TenantAdministration;
using Microsoft.SharePoint.Client;
using PnP.Framework.Diagnostics;
using PnP.Framework.Http;
using PnP.Framework.Provisioning.Model;
using PnP.Framework.Provisioning.Model.Configuration;
using System;
using System.Collections.Generic;
using System.Linq;

namespace PnP.Framework.Provisioning.ObjectHandlers
{
    internal class SiteToTemplateConversion
    {
        /// <summary>
        /// Actual implementation of extracting configuration from existing site.
        /// </summary>
        /// <param name="web"></param>
        /// <param name="creationInfo"></param>
        /// <returns></returns>
        internal static ProvisioningTemplate GetRemoteTemplate(Web web, ProvisioningTemplateCreationInformation creationInfo)
        {

            using (var scope = new PnPMonitoredScope(CoreResources.Provisioning_ObjectHandlers_Extraction))
            {

                web.Context.DisableReturnValueCache = true;
                ProvisioningProgressDelegate progressDelegate = null;
                ProvisioningMessagesDelegate messagesDelegate = null;
                if (creationInfo != null)
                {
                    if (creationInfo.BaseTemplate != null)
                    {
                        scope.LogDebug(CoreResources.SiteToTemplateConversion_Base_template_available___0_, creationInfo.BaseTemplate.Id);
                    }
                    progressDelegate = creationInfo.ProgressDelegate;
                    if (creationInfo.ProgressDelegate != null)
                    {
                        scope.LogDebug(CoreResources.SiteToTemplateConversion_ProgressDelegate_registered);
                    }
                    messagesDelegate = creationInfo.MessagesDelegate;
                    if (creationInfo.MessagesDelegate != null)
                    {
                        scope.LogDebug(CoreResources.SiteToTemplateConversion_MessagesDelegate_registered);
                    }
                    if (creationInfo.IncludeAllTermGroups)
                    {
                        scope.LogDebug(CoreResources.SiteToTemplateConversion_IncludeAllTermGroups_is_set_to_true);
                    }
                    if (creationInfo.IncludeSiteCollectionTermGroup)
                    {
                        scope.LogDebug(CoreResources.SiteToTemplateConversion_IncludeSiteCollectionTermGroup_is_set_to_true);
                    }
                    if (creationInfo.PersistBrandingFiles)
                    {
                        scope.LogDebug(CoreResources.SiteToTemplateConversion_PersistBrandingFiles_is_set_to_true);
                    }
                }
                else
                {
                    // When no provisioning info was passed then we want to execute all handlers
                    creationInfo = new ProvisioningTemplateCreationInformation(web)
                    {
                        HandlersToProcess = Handlers.All
                    };
                }

                // Create empty object
                ProvisioningTemplate template = new ProvisioningTemplate
                {

                    // Hookup connector, is handy when the generated template object is used to apply to another site
                    Connector = creationInfo.FileConnector
                };

                List<ObjectHandlerBase> objectHandlers = new List<ObjectHandlerBase>();

                if (creationInfo.HandlersToProcess.HasFlag(Handlers.RegionalSettings)) objectHandlers.Add(new ObjectRegionalSettings());
                if (creationInfo.HandlersToProcess.HasFlag(Handlers.SupportedUILanguages)) objectHandlers.Add(new ObjectSupportedUILanguages());
                if (creationInfo.HandlersToProcess.HasFlag(Handlers.AuditSettings)) objectHandlers.Add(new ObjectAuditSettings());
                if (creationInfo.HandlersToProcess.HasFlag(Handlers.SitePolicy)) objectHandlers.Add(new ObjectSitePolicy());
                if (creationInfo.HandlersToProcess.HasFlag(Handlers.SiteSecurity)) objectHandlers.Add(new ObjectSiteSecurity());
                if (creationInfo.HandlersToProcess.HasFlag(Handlers.TermGroups)) objectHandlers.Add(new ObjectTermGroups());
                if (creationInfo.HandlersToProcess.HasFlag(Handlers.Fields)) objectHandlers.Add(new ObjectField(FieldAndListProvisioningStepHelper.Step.Export));
                if (creationInfo.HandlersToProcess.HasFlag(Handlers.ContentTypes)) objectHandlers.Add(new ObjectContentType(FieldAndListProvisioningStepHelper.Step.Export));
                if (creationInfo.HandlersToProcess.HasFlag(Handlers.Lists)) objectHandlers.Add(new ObjectListInstance(FieldAndListProvisioningStepHelper.Step.Export));
                if (creationInfo.HandlersToProcess.HasFlag(Handlers.Lists)) objectHandlers.Add(new ObjectListInstanceDataRows());
                if (creationInfo.HandlersToProcess.HasFlag(Handlers.CustomActions)) objectHandlers.Add(new ObjectCustomActions());
                if (creationInfo.HandlersToProcess.HasFlag(Handlers.Features)) objectHandlers.Add(new ObjectFeatures());
                if (creationInfo.HandlersToProcess.HasFlag(Handlers.ComposedLook)) objectHandlers.Add(new ObjectComposedLook());
                if (creationInfo.HandlersToProcess.HasFlag(Handlers.SearchSettings)) objectHandlers.Add(new ObjectSearchSettings());
                if (creationInfo.HandlersToProcess.HasFlag(Handlers.Files)) objectHandlers.Add(new ObjectFiles());
                if (creationInfo.HandlersToProcess.HasFlag(Handlers.Pages)) objectHandlers.Add(new ObjectPages());
                if (creationInfo.HandlersToProcess.HasFlag(Handlers.PageContents)) objectHandlers.Add(new ObjectPageContents());
                if (creationInfo.HandlersToProcess.HasFlag(Handlers.PageContents)) objectHandlers.Add(new ObjectClientSidePageContents());
                if (creationInfo.HandlersToProcess.HasFlag(Handlers.SiteHeader)) objectHandlers.Add(new ObjectSiteHeaderSettings());
                if (creationInfo.HandlersToProcess.HasFlag(Handlers.SiteFooter)) objectHandlers.Add(new ObjectSiteFooterSettings());
                if (creationInfo.HandlersToProcess.HasFlag(Handlers.PropertyBagEntries)) objectHandlers.Add(new ObjectPropertyBagEntry());
                if (creationInfo.HandlersToProcess.HasFlag(Handlers.Publishing)) objectHandlers.Add(new ObjectPublishing());
                if (creationInfo.HandlersToProcess.HasFlag(Handlers.Workflows)) objectHandlers.Add(new ObjectWorkflows());
                if (creationInfo.HandlersToProcess.HasFlag(Handlers.WebSettings)) objectHandlers.Add(new ObjectWebSettings());
                if (creationInfo.HandlersToProcess.HasFlag(Handlers.SiteSettings)) objectHandlers.Add(new ObjectSiteSettings());
                if (creationInfo.HandlersToProcess.HasFlag(Handlers.Theme)) objectHandlers.Add(new ObjectTheme());
                if (creationInfo.HandlersToProcess.HasFlag(Handlers.Navigation)) objectHandlers.Add(new ObjectNavigation());
                if (creationInfo.HandlersToProcess.HasFlag(Handlers.ImageRenditions)) objectHandlers.Add(new ObjectImageRenditions());
                if (creationInfo.HandlersToProcess.HasFlag(Handlers.SyntexModels)) objectHandlers.Add(new ObjectSyntexModels());
                objectHandlers.Add(new ObjectLocalization()); // Always add this one, check is done in the handler
                if (creationInfo.HandlersToProcess.HasFlag(Handlers.Tenant)) objectHandlers.Add(new ObjectTenant());
                if (creationInfo.HandlersToProcess.HasFlag(Handlers.ApplicationLifecycleManagement)) objectHandlers.Add(new ObjectApplicationLifecycleManagement());
                if (creationInfo.HandlersToProcess.HasFlag(Handlers.ExtensibilityProviders)) objectHandlers.Add(new ObjectExtensibilityHandlers());

                objectHandlers.Add(new ObjectRetrieveTemplateInfo());

                int step = 1;

                var count = objectHandlers.Count(o => o.ReportProgress && o.WillExtract(web, template, creationInfo));

                web.EnsureProperty(w => w.Url);

                foreach (var handler in objectHandlers)
                {
                    if (handler.WillExtract(web, template, creationInfo))
                    {
                        if (messagesDelegate != null)
                        {
                            handler.MessagesDelegate = messagesDelegate;
                        }
                        if (handler.ReportProgress && progressDelegate != null)
                        {
                            progressDelegate(handler.Name, step, count);
                            step++;
                        }

                        using (var handlerContext = web.Context.Clone(web.Url))
                        {
                            template = handler.ExtractObjects(handlerContext.Web, template, creationInfo);
                        }
                    }
                }

                return template;
            }
        }

        internal void ApplyTenantTemplate(Tenant tenant, PnP.Framework.Provisioning.Model.ProvisioningHierarchy hierarchy, string sequenceId, ApplyConfiguration configuration)
        {
            using (var scope = new PnPMonitoredScope(CoreResources.Provisioning_ObjectHandlers_Provisioning))
            {
                ProvisioningProgressDelegate progressDelegate = null;
                ProvisioningMessagesDelegate messagesDelegate = null;
                if (configuration == null)
                {
                    // When no provisioning info was passed then we want to execute all handlers
                    configuration = new ApplyConfiguration();
                }
                else
                {
                    progressDelegate = configuration.ProgressDelegate;
                    if (configuration.ProgressDelegate != null)
                    {
                        scope.LogInfo(CoreResources.SiteToTemplateConversion_ProgressDelegate_registered);
                    }
                    messagesDelegate = configuration.MessagesDelegate;
                    if (configuration.MessagesDelegate != null)
                    {
                        scope.LogInfo(CoreResources.SiteToTemplateConversion_MessagesDelegate_registered);
                    }
                }

                Log.Debug(Constants.LOGGING_SOURCE, $"Attaching object handlers");

                List<ObjectHierarchyHandlerBase> objectHandlers = new List<ObjectHierarchyHandlerBase>
                {
                    new ObjectHierarchyTenant(),
                    new ObjectHierarchySequenceTermGroups(),
                    new ObjectHierarchySequenceSites(),
                    new ObjectTeams(),
                    new ObjectAzureActiveDirectory(),
                };

                var count = objectHandlers.Count(o => o.ReportProgress && o.WillProvision(tenant, hierarchy, sequenceId, configuration)) + 1;

                progressDelegate?.Invoke("Initializing engine", 1, count); // handlers + initializing message)

                int step = 2;

                TokenParser sequenceTokenParser = new TokenParser(tenant, hierarchy);

                CallWebHooks(hierarchy.Templates.FirstOrDefault(), sequenceTokenParser,
                    ProvisioningTemplateWebhookKind.ProvisioningStarted);

                foreach (var handler in objectHandlers)
                {
                    if (handler.WillProvision(tenant, hierarchy, sequenceId, configuration))
                    {
                        if (messagesDelegate != null)
                        {
                            handler.MessagesDelegate = messagesDelegate;
                        }
                        if (handler.ReportProgress && progressDelegate != null)
                        {
                            progressDelegate(handler.Name, step, count);
                            step++;
                        }
                        try
                        {
                            sequenceTokenParser = handler.ProvisionObjects(tenant, hierarchy, sequenceId, sequenceTokenParser, configuration);
                        }
                        catch (Exception ex)
                        {
                            CallWebHooks(hierarchy.Templates.FirstOrDefault(), sequenceTokenParser,
                                ProvisioningTemplateWebhookKind.ProvisioningExceptionOccurred, handler.Name, ex);
                            throw;
                        }
                    }
                }

                CallWebHooks(hierarchy.Templates.FirstOrDefault(), sequenceTokenParser,
                    ProvisioningTemplateWebhookKind.ProvisioningCompleted);

            }
        }

        internal static ProvisioningHierarchy GetTenantTemplate(Tenant tenant, ExtractConfiguration configuration = null)
        {
            if (configuration == null)
            {
                configuration = new ExtractConfiguration();
            }

            using (var scope = new PnPMonitoredScope(CoreResources.Provisioning_ObjectHandlers_Extraction))
            {

                ProvisioningHierarchy tenantTemplate = new ProvisioningHierarchy
                {
                    Connector = configuration.FileConnector
                };

                List<ObjectHierarchyHandlerBase> objectHandlers = new List<ObjectHierarchyHandlerBase>();

                if (configuration.Tenant.Sequence != null) objectHandlers.Add(new ObjectHierarchySequenceSites()); // always build up the sequence
                if (configuration.Tenant.Teams != null) objectHandlers.Add(new ObjectTeams());

                int step = 1;

                var count = objectHandlers.Count(o => o.ReportProgress && o.WillExtract(tenant, tenantTemplate, null, configuration));

                foreach (var handler in objectHandlers)
                {
                    if (handler.WillExtract(tenant, tenantTemplate, null, null))
                    {
                        if (configuration.MessagesDelegate != null)
                        {
                            handler.MessagesDelegate = (message, type) =>
                            {
                                configuration.MessagesDelegate(message, type);
                            };
                        }
                        if (handler.ReportProgress && configuration.ProgressDelegate != null)
                        {
                            configuration.ProgressDelegate(handler.Name, step, count);
                            step++;
                        }

                        tenantTemplate = handler.ExtractObjects(tenant, tenantTemplate, configuration);
                    }
                }

                return tenantTemplate;
            }
        }

        /// <summary>
        /// Actual implementation of the apply templates
        /// </summary>
        /// <param name="web"></param>
        /// <param name="template"></param>
        /// <param name="provisioningInfo"></param>
        /// <param name="calledFromHierarchy"></param>
        /// <param name="tokenParser"></param>
        internal void ApplyRemoteTemplate(Web web, ProvisioningTemplate template, ProvisioningTemplateApplyingInformation provisioningInfo, bool calledFromHierarchy = false, TokenParser tokenParser = null)
        {
            using (var scope = new PnPMonitoredScope(CoreResources.Provisioning_ObjectHandlers_Provisioning))
            {

                web.Context.DisableReturnValueCache = true;

                ProvisioningProgressDelegate progressDelegate = null;
                ProvisioningMessagesDelegate messagesDelegate = null;
                ProvisioningSiteProvisionedDelegate siteProvisionedDelegate = null;
                if (provisioningInfo != null)
                {
                    if (provisioningInfo.OverwriteSystemPropertyBagValues == true)
                    {
                        scope.LogInfo(CoreResources.SiteToTemplateConversion_ApplyRemoteTemplate_OverwriteSystemPropertyBagValues_is_to_true);
                    }
                    progressDelegate = provisioningInfo.ProgressDelegate;
                    if (provisioningInfo.ProgressDelegate != null)
                    {
                        scope.LogInfo(CoreResources.SiteToTemplateConversion_ProgressDelegate_registered);
                    }
                    messagesDelegate = provisioningInfo.MessagesDelegate;
                    if (provisioningInfo.MessagesDelegate != null)
                    {
                        scope.LogInfo(CoreResources.SiteToTemplateConversion_MessagesDelegate_registered);
                    }
                    siteProvisionedDelegate = provisioningInfo.SiteProvisionedDelegate;
                }
                else
                {
                    // When no provisioning info was passed then we want to execute all handlers
                    provisioningInfo = new ProvisioningTemplateApplyingInformation
                    {
                        HandlersToProcess = Handlers.All
                    };
                }

                // Check if scope is present and if so, matches the current site. When scope was not set the returned value will be ProvisioningTemplateScope.Undefined
                if (template.Scope == ProvisioningTemplateScope.RootSite)
                {
                    if (web.IsSubSite())
                    {
                        scope.LogError(CoreResources.SiteToTemplateConversion_ScopeOfTemplateDoesNotMatchTarget);
                        throw new Exception(CoreResources.SiteToTemplateConversion_ScopeOfTemplateDoesNotMatchTarget);
                    }
                }
                var currentCultureInfoValue = System.Threading.Thread.CurrentThread.CurrentCulture;
                if (!string.IsNullOrEmpty(template.TemplateCultureInfo))
                {
                    if (int.TryParse(template.TemplateCultureInfo, out var cultureInfoValue))
                    {
                        System.Threading.Thread.CurrentThread.CurrentCulture = new System.Globalization.CultureInfo(cultureInfoValue);
                    }
                    else
                    {
                        System.Threading.Thread.CurrentThread.CurrentCulture = new System.Globalization.CultureInfo(template.TemplateCultureInfo);
                    }
                }

                // Check if the target site shares the same base template with the template's source site
                var targetSiteTemplateId = web.GetBaseTemplateId();
                if (!String.IsNullOrEmpty(targetSiteTemplateId) && !String.IsNullOrEmpty(template.BaseSiteTemplate))
                {
                    if (!targetSiteTemplateId.Equals(template.BaseSiteTemplate, StringComparison.InvariantCultureIgnoreCase))
                    {
                        var templatesNotMatchingWarning = String.Format(CoreResources.Provisioning_Asymmetric_Base_Templates, template.BaseSiteTemplate, targetSiteTemplateId);
                        scope.LogWarning(templatesNotMatchingWarning);
                        messagesDelegate?.Invoke(templatesNotMatchingWarning, ProvisioningMessageType.Warning);
                    }
                }

                // Always ensure the Url property is loaded. In the tokens we need this and we don't want to call ExecuteQuery as this can
                // impact delta scenarions (calling ExecuteQuery before the planned update is called)
                web.EnsureProperty(w => w.Url);


                List<ObjectHandlerBase> objectHandlers = new List<ObjectHandlerBase>();

                if (provisioningInfo.HandlersToProcess.HasFlag(Handlers.RegionalSettings)) objectHandlers.Add(new ObjectRegionalSettings());
                if (provisioningInfo.HandlersToProcess.HasFlag(Handlers.SupportedUILanguages)) objectHandlers.Add(new ObjectSupportedUILanguages());
                if (provisioningInfo.HandlersToProcess.HasFlag(Handlers.AuditSettings)) objectHandlers.Add(new ObjectAuditSettings());
                if (provisioningInfo.HandlersToProcess.HasFlag(Handlers.SitePolicy)) objectHandlers.Add(new ObjectSitePolicy());
                if (provisioningInfo.HandlersToProcess.HasFlag(Handlers.SiteSecurity)) objectHandlers.Add(new ObjectSiteSecurity());
                if (provisioningInfo.HandlersToProcess.HasFlag(Handlers.Features)) objectHandlers.Add(new ObjectFeatures());
                if (provisioningInfo.HandlersToProcess.HasFlag(Handlers.TermGroups)) objectHandlers.Add(new ObjectTermGroups());

                // Process 3 times these providers to handle proper ordering of artefact creation when dealing with lookup fields

                // 1st. create fields, content and list without lookup fields
                if (provisioningInfo.HandlersToProcess.HasFlag(Handlers.Fields) || provisioningInfo.HandlersToProcess.HasFlag(Handlers.Lists)) objectHandlers.Add(new ObjectField(FieldAndListProvisioningStepHelper.Step.ListAndStandardFields));
                if (provisioningInfo.HandlersToProcess.HasFlag(Handlers.ContentTypes)) objectHandlers.Add(new ObjectContentType(FieldAndListProvisioningStepHelper.Step.ListAndStandardFields));
                if (provisioningInfo.HandlersToProcess.HasFlag(Handlers.Lists)) objectHandlers.Add(new ObjectListInstance(FieldAndListProvisioningStepHelper.Step.ListAndStandardFields));

                // 2nd. create lookup fields (which requires lists to be present
                if (provisioningInfo.HandlersToProcess.HasFlag(Handlers.Fields) || provisioningInfo.HandlersToProcess.HasFlag(Handlers.Lists)) objectHandlers.Add(new ObjectField(FieldAndListProvisioningStepHelper.Step.LookupFields));
                if (provisioningInfo.HandlersToProcess.HasFlag(Handlers.ContentTypes)) objectHandlers.Add(new ObjectContentType(FieldAndListProvisioningStepHelper.Step.LookupFields));
                if (provisioningInfo.HandlersToProcess.HasFlag(Handlers.Lists)) objectHandlers.Add(new ObjectListInstance(FieldAndListProvisioningStepHelper.Step.LookupFields));

                if (provisioningInfo.HandlersToProcess.HasFlag(Handlers.Files)) objectHandlers.Add(new ObjectFiles());

                // 3rd. Create remaining objects in lists (views, user custom actions, ...)
                if (provisioningInfo.HandlersToProcess.HasFlag(Handlers.Lists)) objectHandlers.Add(new ObjectListInstance(FieldAndListProvisioningStepHelper.Step.ListSettings));

                if (provisioningInfo.HandlersToProcess.HasFlag(Handlers.Fields) || provisioningInfo.HandlersToProcess.HasFlag(Handlers.Lists)) objectHandlers.Add(new ObjectListInstanceDataRows());
                if (provisioningInfo.HandlersToProcess.HasFlag(Handlers.Workflows)) objectHandlers.Add(new ObjectWorkflows());
                if (provisioningInfo.HandlersToProcess.HasFlag(Handlers.Pages)) objectHandlers.Add(new ObjectPages());
                if (provisioningInfo.HandlersToProcess.HasFlag(Handlers.PageContents)) objectHandlers.Add(new ObjectPageContents());
                if (!calledFromHierarchy && provisioningInfo.HandlersToProcess.HasFlag(Handlers.Tenant)) objectHandlers.Add(new ObjectTenant());
                if (provisioningInfo.HandlersToProcess.HasFlag(Handlers.ApplicationLifecycleManagement)) objectHandlers.Add(new ObjectApplicationLifecycleManagement());
                if (provisioningInfo.HandlersToProcess.HasFlag(Handlers.Pages)) objectHandlers.Add(new ObjectClientSidePages());
                if (provisioningInfo.HandlersToProcess.HasFlag(Handlers.SiteHeader)) objectHandlers.Add(new ObjectSiteHeaderSettings());
                if (provisioningInfo.HandlersToProcess.HasFlag(Handlers.SiteFooter)) objectHandlers.Add(new ObjectSiteFooterSettings());
                if (provisioningInfo.HandlersToProcess.HasFlag(Handlers.CustomActions)) objectHandlers.Add(new ObjectCustomActions());
                if (provisioningInfo.HandlersToProcess.HasFlag(Handlers.Publishing)) objectHandlers.Add(new ObjectPublishing());
                if (provisioningInfo.HandlersToProcess.HasFlag(Handlers.ComposedLook)) objectHandlers.Add(new ObjectComposedLook());
                if (provisioningInfo.HandlersToProcess.HasFlag(Handlers.SearchSettings)) objectHandlers.Add(new ObjectSearchSettings());
                if (provisioningInfo.HandlersToProcess.HasFlag(Handlers.PropertyBagEntries)) objectHandlers.Add(new ObjectPropertyBagEntry());
                if (provisioningInfo.HandlersToProcess.HasFlag(Handlers.WebSettings)) objectHandlers.Add(new ObjectWebSettings());
                if (provisioningInfo.HandlersToProcess.HasFlag(Handlers.SiteSettings)) objectHandlers.Add(new ObjectSiteSettings());
                if (provisioningInfo.HandlersToProcess.HasFlag(Handlers.Theme)) objectHandlers.Add(new ObjectTheme());
                if (provisioningInfo.HandlersToProcess.HasFlag(Handlers.Navigation)) objectHandlers.Add(new ObjectNavigation());
                if (provisioningInfo.HandlersToProcess.HasFlag(Handlers.ImageRenditions)) objectHandlers.Add(new ObjectImageRenditions());
                objectHandlers.Add(new ObjectLocalization()); // Always add this one, check is done in the handler
                if (provisioningInfo.HandlersToProcess.HasFlag(Handlers.ExtensibilityProviders)) objectHandlers.Add(new ObjectExtensibilityHandlers());

                // Only persist template information in case this flag is set: this will allow the engine to
                // work with lesser permissions
                if (provisioningInfo.PersistTemplateInfo)
                {
                    objectHandlers.Add(new ObjectPersistTemplateInfo());
                }
                var count = objectHandlers.Count(o => o.ReportProgress && o.WillProvision(web, template, provisioningInfo)) + 1;

                progressDelegate?.Invoke("Initializing engine", 1, count); // handlers + initializing message)
                if (tokenParser == null)
                {
                    tokenParser = new TokenParser(web, template);
                }
                if (provisioningInfo.HandlersToProcess.HasFlag(Handlers.ExtensibilityProviders))
                {
                    var extensibilityHandler = objectHandlers.OfType<ObjectExtensibilityHandlers>().First();
                    extensibilityHandler.AddExtendedTokens(web, template, tokenParser, provisioningInfo);
                }

                int step = 2;

                // Remove potentially unsupported artifacts

                var cleaner = new NoScriptTemplateCleaner(web);
                if (messagesDelegate != null)
                {
                    cleaner.MessagesDelegate = messagesDelegate;
                }
                template = cleaner.CleanUpBeforeProvisioning(template);

                CallWebHooks(template, tokenParser, ProvisioningTemplateWebhookKind.ProvisioningTemplateStarted);

                foreach (var handler in objectHandlers)
                {
                    if (handler.WillProvision(web, template, provisioningInfo))
                    {
                        if (messagesDelegate != null)
                        {
                            handler.MessagesDelegate = messagesDelegate;
                        }
                        if (handler.ReportProgress && progressDelegate != null)
                        {
                            progressDelegate(handler.Name, step, count);
                            step++;
                        }
                        CallWebHooks(template, tokenParser,
                            ProvisioningTemplateWebhookKind.ObjectHandlerProvisioningStarted,
                            handler.InternalName);

                        try
                        {
                            tokenParser = handler.ProvisionObjects(web, template, tokenParser, provisioningInfo);
                        }
                        catch (Exception ex)
                        {
                            CallWebHooks(template, tokenParser, ProvisioningTemplateWebhookKind.ExceptionOccurred,
                                handler.InternalName, ex);
                            throw;
                        }
                        CallWebHooks(template, tokenParser,
                            ProvisioningTemplateWebhookKind.ObjectHandlerProvisioningCompleted,
                            handler.InternalName);
                    }
                }

                // Notify the completed provisioning of the site
                web.EnsureProperties(w => w.Title, w => w.Url);
                siteProvisionedDelegate?.Invoke(web.Title, web.Url);

                CallWebHooks(template, tokenParser, ProvisioningTemplateWebhookKind.ProvisioningTemplateCompleted);

                System.Threading.Thread.CurrentThread.CurrentCulture = currentCultureInfoValue;

            }
        }

        private static void CallWebHooks(ProvisioningTemplate template, TokenParser parser, ProvisioningTemplateWebhookKind kind, String objectHandler = null, Exception exception = null)
        {
            if (template != null)
            {
                using (var scope = new PnPMonitoredScope("ProvisioningTemplate WebHook Call"))
                {
                    var webhooks = new List<ProvisioningWebhookBase>();

                    // Merge the webhooks at template level with those at global level
                    if (template.ProvisioningTemplateWebhooks != null && template.ProvisioningTemplateWebhooks.Any())
                    {
                        webhooks.AddRange(template.ProvisioningTemplateWebhooks);
                    }
                    if (template.ParentHierarchy?.ProvisioningWebhooks != null && template.ParentHierarchy.ProvisioningWebhooks.Any())
                    {
                        webhooks.AddRange(template.ParentHierarchy.ProvisioningWebhooks);
                    }

                    // If there is any webhook
                    if (webhooks.Count > 0)
                    {
                        foreach (var webhook in webhooks.Where(w => w.Kind == kind))
                        {
                            WebhookSender.InvokeWebhook(webhook, PnPHttpClient.Instance.GetHttpClient(), kind, parser, objectHandler, exception, scope);
                        }
                    }
                }
            }
        }
    }
}