﻿using Newtonsoft.Json;
using PnP.Framework.Graph.Model;
using PnP.Framework.Utilities;
using System;
using System.Collections.Generic;
using System.Linq;

namespace PnP.Framework.Graph
{
    /// <summary>
    /// Utility class for managing Site Classifications settings on the target tenant using Graph.
    /// </summary>
    public static class SiteClassificationsUtility
    {
        /// <summary>
        /// Disables Site Classifications settings for the target tenant
        /// </summary>
        /// <param name="accessToken">The OAuth accessToken for Microsoft Graph with Azure AD</param>
        /// <param name="azureEnvironment">Defines the Azure Cloud Deployment. This is used to determine the MS Graph EndPoint to call which differs per Azure Cloud deployments. Defaults to Production (graph.microsoft.com).</param>
        public static void DisableSiteClassifications(string accessToken, AzureEnvironment azureEnvironment = AzureEnvironment.Production)
        {
            if (string.IsNullOrEmpty(accessToken))
            {
                throw new ArgumentException("Specify a valid accesstoken", nameof(accessToken));
            }
            // GET https://graph.microsoft.com/beta/settings
            string directorySettingsUrl = $"{GraphHttpClient.GetGraphEndPointUrl(azureEnvironment,true)}settings";
            var directorySettingsJson = GraphHttpClient.MakeGetRequestForString(directorySettingsUrl, accessToken);
            var directorySettings = JsonConvert.DeserializeObject<DirectorySettingTemplates>(directorySettingsJson);

            // Retrieve the setinngs for "Group.Unified"
            var unifiedGroupSetting = directorySettings.Templates.FirstOrDefault(t => t.DisplayName == "Group.Unified");

            if (unifiedGroupSetting != null)
            {
                // DELETE https://graph.microsoft.com/beta/settings
                string deleteDirectorySettingUrl = $"{GraphHttpClient.GetGraphEndPointUrl(azureEnvironment, true)}settings/{unifiedGroupSetting.Id}";
                GraphHttpClient.MakeDeleteRequest(
                    deleteDirectorySettingUrl,
                    accessToken: accessToken);
            }
            else
            {
                throw new ApplicationException("Missing DirectorySettingTemplate for \"Group.Unified\"");
            }
        }

        /// <summary>
        /// Enables Site Classifications for the target tenant 
        /// </summary>
        /// <param name="accessToken">The OAuth accessToken for Microsoft Graph with Azure AD</param>
        /// <param name="siteClassificationsSettings">The site classifications settings to apply.</param>
        public static void EnableSiteClassifications(string accessToken, SiteClassificationsSettings siteClassificationsSettings)
        {
            if (string.IsNullOrEmpty(accessToken))
            {
                throw new ArgumentException("Specify a valid accesstoken", nameof(accessToken));
            }
            if (siteClassificationsSettings == null)
            {
                throw new ArgumentException(nameof(siteClassificationsSettings));
            }
            EnableSiteClassifications(accessToken, siteClassificationsSettings.Classifications, siteClassificationsSettings.DefaultClassification, siteClassificationsSettings.DefaultClassification);
        }

        /// <summary>
        /// Enables Site Classifications for the target tenant 
        /// </summary>
        /// <param name="accessToken">The OAuth accessToken for Microsoft Graph with Azure AD</param>
        /// <param name="classificationsList">The list of classification values</param>
        /// <param name="defaultClassification">The default classification</param>
        /// <param name="usageGuidelinesUrl">The URL of a guidance page</param>
        /// <param name="azureEnvironment">Defines the Azure Cloud Deployment. This is used to determine the MS Graph EndPoint to call which differs per Azure Cloud deployments. Defaults to Production (graph.microsoft.com).</param>
        public static void EnableSiteClassifications(string accessToken, IEnumerable<string> classificationsList, string defaultClassification = "", string usageGuidelinesUrl = "", AzureEnvironment azureEnvironment = AzureEnvironment.Production)
        {
            if (string.IsNullOrEmpty(accessToken))
            {
                throw new ArgumentException("Specify a valid accesstoken", nameof(accessToken));
            }
            if (classificationsList == null || !classificationsList.Any())
            {
                throw new ArgumentException("Specify one or more classifications", nameof(classificationsList));
            }
            if (usageGuidelinesUrl == null)
            {
                throw new ArgumentException("Specify a valid URL or an empty string to not set this value", nameof(usageGuidelinesUrl));
            }
            if (!classificationsList.Contains(defaultClassification))
            {
                throw new ArgumentException("The default classification specified is not available in the list of specified classifications", nameof(defaultClassification));
            }

            // GET https://graph.microsoft.com/beta/directorySettingTemplates
            string directorySettingTemplatesUrl = $"{GraphHttpClient.GetGraphEndPointUrl(azureEnvironment,true)}directorySettingTemplates";
            var directorySettingTemplatesJson = GraphHttpClient.MakeGetRequestForString(directorySettingTemplatesUrl, accessToken);
            var directorySettingTemplates = JsonConvert.DeserializeObject<DirectorySettingTemplates>(directorySettingTemplatesJson);

            // Retrieve the setinngs for "Group.Unified"
            var unifiedGroupSetting = directorySettingTemplates.Templates.FirstOrDefault(t => t.DisplayName == "Group.Unified");

            if (unifiedGroupSetting != null)
            {
                var directorySettingValues = new Dictionary<string, string>();
                foreach (var v in unifiedGroupSetting.SettingValues)
                {
                    switch (v.Name)
                    {
                        case "UsageGuidelinesUrl":
                            directorySettingValues.Add(v.Name, usageGuidelinesUrl);
                            break;
                        case "ClassificationList":
                            directorySettingValues.Add(v.Name, classificationsList.Aggregate((s, i) => s + ", " + i));
                            break;
                        case "DefaultClassification":
                            directorySettingValues.Add(v.Name, defaultClassification);
                            break;
                        default:
                            directorySettingValues.Add(v.Name, v.DefaultValue);
                            break;
                    }
                }

                // POST https://graph.microsoft.com/beta/settings
                string newDirectorySettingUrl = $"{GraphHttpClient.GetGraphEndPointUrl(azureEnvironment,true)}settings";
                var newDirectorySettingResult = GraphHttpClient.MakePostRequestForString(
                    newDirectorySettingUrl,
                    content: new
                    {
                        templateId = unifiedGroupSetting.Id,
                        values = from v in directorySettingValues select new { name = v.Key, value = v.Value },
                    },
                    contentType: HttpHelper.JsonContentType,
                    accessToken: accessToken);
            }
            else
            {
                throw new ApplicationException("Missing DirectorySettingTemplate for \"Group.Unified\"");
            }
        }

        /// <summary>
        /// Enables Site Classifications for the target tenant 
        /// </summary>
        /// <param name="accessToken">The OAuth accessToken for Microsoft Graph with Azure AD</param>
        /// <param name="azureEnvironment">Defines the Azure Cloud Deployment. This is used to determine the MS Graph EndPoint to call which differs per Azure Cloud deployments. Defaults to Production (graph.microsoft.com).</param>
        /// <returns>The list of Site Classification values</returns>
        public static SiteClassificationsSettings GetSiteClassificationsSettings(string accessToken, AzureEnvironment azureEnvironment = AzureEnvironment.Production)
        {
            if (string.IsNullOrEmpty(accessToken))
            {
                throw new ArgumentException("Specify a valid accesstoken", nameof(accessToken));
            }
            // GET https://graph.microsoft.com/beta/directorySettingTemplates
            string directorySettingsUrl = $"{GraphHttpClient.GetGraphEndPointUrl(azureEnvironment,true)}settings";
            var directorySettingsJson = GraphHttpClient.MakeGetRequestForString(directorySettingsUrl, accessToken);
            var directorySettings = JsonConvert.DeserializeObject<DirectorySettingTemplates>(directorySettingsJson);

            // Retrieve the setinngs for "Group.Unified"
            var unifiedGroupSetting = directorySettings.Templates.FirstOrDefault(t => t.DisplayName == "Group.Unified");

            if (unifiedGroupSetting != null)
            {
                var siteClassificationsSettings = new SiteClassificationsSettings();
                var classificationList = unifiedGroupSetting.SettingValues.FirstOrDefault(v => v.Name == "ClassificationList");
                if (classificationList != null)
                {
                    siteClassificationsSettings.Classifications = classificationList.Value.Split(new char[] { ',' }, StringSplitOptions.RemoveEmptyEntries).ToList();
                }
                var guidanceUrl = unifiedGroupSetting.SettingValues.First(v => v.Name == "UsageGuidelinesUrl");
                if (guidanceUrl != null)
                {
                    siteClassificationsSettings.UsageGuidelinesUrl = guidanceUrl.Value;
                }
                var defaultClassification = unifiedGroupSetting.SettingValues.First(v => v.Name == "DefaultClassification");
                if (defaultClassification != null)
                {
                    siteClassificationsSettings.DefaultClassification = defaultClassification.Value;
                }
                return siteClassificationsSettings;
            }
            else
            {
                throw new ApplicationException("Missing DirectorySettingTemplate for \"Group.Unified\"");
            }
        }

        /// <summary>
        /// Updates Site Classifications settings for the target tenant
        /// </summary>
        /// <param name="accessToken">The OAuth accessToken for Microsoft Graph with Azure AD</param>
        /// <param name="siteClassificationsSettings">The site classifications settings to apply.</param>
        public static void UpdateSiteClassificationsSettings(string accessToken, SiteClassificationsSettings siteClassificationsSettings)
        {
            if (string.IsNullOrEmpty(accessToken))
            {
                throw new ArgumentException("Specify a valid accesstoken", nameof(accessToken));
            }
            if (siteClassificationsSettings == null)
            {
                throw new ArgumentException("Specify a valid Site Classification Settings object", nameof(siteClassificationsSettings));
            }
            UpdateSiteClassificationsSettings(accessToken, siteClassificationsSettings.Classifications, siteClassificationsSettings.DefaultClassification, siteClassificationsSettings.UsageGuidelinesUrl);
        }

        /// <summary>
        /// Updates Site Classifications settings for the target tenant
        /// </summary>
        /// <param name="accessToken">The OAuth accessToken for Microsoft Graph with Azure AD</param>
        /// <param name="classificationsList">The list of classification values</param>
        /// <param name="defaultClassification">The default classification</param>
        /// <param name="usageGuidelinesUrl">The URL of a guidance page</param>
        /// <param name="azureEnvironment">Defines the Azure Cloud Deployment. This is used to determine the MS Graph EndPoint to call which differs per Azure Cloud deployments. Defaults to Production (graph.microsoft.com).</param>
        public static void UpdateSiteClassificationsSettings(string accessToken, IEnumerable<string> classificationsList = null, string defaultClassification = "", string usageGuidelinesUrl = "", AzureEnvironment azureEnvironment = AzureEnvironment.Production)
        {
            if (string.IsNullOrEmpty(accessToken))
            {
                throw new ArgumentException("Specify a valid accesstoken", nameof(accessToken));
            }
            // GET https://graph.microsoft.com/beta/settings
            string directorySettingsUrl = $"{GraphHttpClient.GetGraphEndPointUrl(azureEnvironment,true)}settings";
            var directorySettingsJson = GraphHttpClient.MakeGetRequestForString(directorySettingsUrl, accessToken);
            var directorySettings = JsonConvert.DeserializeObject<DirectorySettingTemplates>(directorySettingsJson);

            // Retrieve the setinngs for "Group.Unified"
            var unifiedGroupSetting = directorySettings.Templates.FirstOrDefault(t => t.DisplayName == "Group.Unified");

            if (unifiedGroupSetting != null)
            {
                foreach (var v in unifiedGroupSetting.SettingValues)
                {
                    switch (v.Name)
                    {
                        case "UsageGuidelinesUrl":
                            if (usageGuidelinesUrl != null)
                            {
                                v.Value = usageGuidelinesUrl;
                            }
                            break;
                        case "ClassificationList":
                            if (classificationsList != null && classificationsList.Any())
                            {
                                v.Value = classificationsList.Aggregate((s, i) => s + ", " + i);
                            }
                            break;
                        case "DefaultClassification":
                            if (usageGuidelinesUrl != null)
                            {
                                v.Value = defaultClassification;
                            }
                            break;
                        default:
                            break;
                    }
                }

                // PATCH https://graph.microsoft.com/beta/settings
                string updateDirectorySettingUrl = $"{GraphHttpClient.GetGraphEndPointUrl(azureEnvironment, true)}settings/{unifiedGroupSetting.Id}";
                var updateDirectorySettingResult = GraphHttpClient.MakePatchRequestForString(
                    updateDirectorySettingUrl,
                    content: new
                    {
                        templateId = unifiedGroupSetting.Id,
                        values = from v in unifiedGroupSetting.SettingValues select new { name = v.Name, value = v.Value },
                    },
                    contentType: HttpHelper.JsonContentType,
                    accessToken: accessToken);
            }
            else
            {
                throw new ApplicationException("Missing DirectorySetting for \"Group.Unified\"");
            }
        }
    }
}
