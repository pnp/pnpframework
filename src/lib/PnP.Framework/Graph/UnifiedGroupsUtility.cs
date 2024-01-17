﻿using Microsoft.Graph;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using PnP.Framework.Diagnostics;
using PnP.Framework.Entities;
using PnP.Framework.Utilities;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net.Http.Headers;
using System.Text.Json;
using System.Threading.Tasks;

namespace PnP.Framework.Graph
{
    public class GroupExtended : Group
    {
#pragma warning disable CA1819
        [JsonProperty("owners@odata.bind", NullValueHandling = NullValueHandling.Ignore)]
        public string[] OwnersODataBind { get; set; }
        [JsonProperty("members@odata.bind", NullValueHandling = NullValueHandling.Ignore)]
        public string[] MembersODataBind { get; set; }
#pragma warning restore CA1819
    }
    /// <summary>
    /// Class that deals with Unified group CRUD operations.
    /// </summary>
    public static class UnifiedGroupsUtility
    {
        private const int defaultRetryCount = 10;
        private const int defaultDelay = 500;

        /// <summary>
        ///  Creates a new GraphServiceClient instance using a custom PnPHttpProvider
        /// </summary>
        /// <param name="accessToken">The OAuth 2.0 Access Token to configure the HTTP bearer Authorization Header</param>
        /// <param name="retryCount">Number of times to retry the request in case of throttling</param>
        /// <param name="delay">Milliseconds to wait before retrying the request.</param>
        /// <param name="azureEnvironment">Azure environment to use</param>
        /// <returns></returns>
        private static GraphServiceClient CreateGraphClient(String accessToken, int retryCount = defaultRetryCount, int delay = defaultDelay, AzureEnvironment azureEnvironment = AzureEnvironment.Production)
        {
            // Creates a new GraphServiceClient instance using a custom PnPHttpProvider
            // which natively supports retry logic for throttled requests
            // Default are 10 retries with a base delay of 500ms

            var baseUrl = $"https://{AuthenticationManager.GetGraphEndPoint(azureEnvironment)}/v1.0";

            var result = new GraphServiceClient(baseUrl, new DelegateAuthenticationProvider(
                        async (requestMessage) =>
                        {
                            await Task.Run(() =>
                            {
                                if (!String.IsNullOrEmpty(accessToken))
                                {
                                    // Configure the HTTP bearer Authorization Header
                                    requestMessage.Headers.Authorization = new AuthenticationHeaderValue("bearer", accessToken);
                                }
                            });
                        }), new PnPHttpProvider(retryCount, delay));

            return (result);
        }

        /// <summary>
        /// Returns the URL of the Modern SharePoint Site backing an Office 365 Group (i.e. Unified Group)
        /// </summary>
        /// <param name="groupId">The ID of the Office 365 Group</param>
        /// <param name="accessToken">The OAuth 2.0 Access Token to use for invoking the Microsoft Graph</param>
        /// <param name="retryCount">Number of times to retry the request in case of throttling</param>
        /// <param name="delay">Milliseconds to wait before retrying the request. The delay will be increased (doubled) every retry</param>
        /// <param name="azureEnvironment">Defines the Azure Cloud Deployment. This is used to determine the MS Graph EndPoint to call which differs per Azure Cloud deployments. Defaults to Production (graph.microsoft.com).</param>
        /// <returns>The URL of the modern site backing the Office 365 Group</returns>
        public static string GetUnifiedGroupSiteUrl(string groupId, string accessToken, int retryCount = 10, int delay = 500, AzureEnvironment azureEnvironment = AzureEnvironment.Production)
        {
            if (string.IsNullOrEmpty(groupId))
            {
                throw new ArgumentNullException(nameof(groupId));
            }

            if (string.IsNullOrEmpty(accessToken))
            {
                throw new ArgumentNullException(nameof(accessToken));
            }
            string result;
            try
            {
                string requestUrl = $"{GraphHttpClient.GetGraphEndPointUrl(azureEnvironment)}groups/{groupId}/sites/root?$select=webUrl";
                var response = JToken.Parse(HttpHelper.MakeGetRequestForString(requestUrl, accessToken, retryCount: retryCount, delay: delay));

                result = Convert.ToString(response["webUrl"]);
            }
            catch (ServiceException ex)
            {
                Log.Error(Constants.LOGGING_SOURCE, CoreResources.GraphExtensions_ErrorOccured, ex.Error.Message);
                throw;
            }
            return (result);
        }

        /// <summary>
        /// Creates a new Office 365 Group (i.e. Unified Group) with its backing Modern SharePoint Site
        /// </summary>
        /// <param name="displayName">The Display Name for the Office 365 Group</param>
        /// <param name="description">The Description for the Office 365 Group</param>
        /// <param name="mailNickname">The Mail Nickname for the Office 365 Group</param>
        /// <param name="accessToken">The OAuth 2.0 Access Token to use for invoking the Microsoft Graph</param>
        /// <param name="owners">A list of UPNs for group owners, if any</param>
        /// <param name="members">A list of UPNs for group members, if any</param>
        /// <param name="groupLogo">The binary stream of the logo for the Office 365 Group</param>
        /// <param name="isPrivate">Defines whether the group will be private or public, optional with default false (i.e. public)</param>
        /// <param name="createTeam">Defines whether to create MS Teams team associated with the group</param>
        /// <param name="retryCount">Number of times to retry the request in case of throttling</param>
        /// <param name="delay">Milliseconds to wait before retrying the request. The delay will be increased (doubled) every retry</param>
        /// <param name="azureEnvironment">Defines the Azure Cloud Deployment. This is used to determine the MS Graph EndPoint to call which differs per Azure Cloud deployments. Defaults to Production (graph.microsoft.com).</param>
        /// <param name="preferredDataLocation">Defines the codes of geographies in which there is Office 365 presence. Used for multi-geo enabled tenants. List with available geographies is available at https://docs.microsoft.com/office365/enterprise/multi-geo-add-group-with-pdl#geo-location-codes.</param>
        /// <param name="assignedLabels">AIP Labels which should be applied to the group (does not work for App-Only)</param>
        /// <param name="welcomeEmailDisabled">Option to prevent sending of default welcome emails to new members.</param>
        /// <returns>The just created Office 365 Group</returns>
        public static UnifiedGroupEntity CreateUnifiedGroup(string displayName, string description, string mailNickname,
            string accessToken, string[] owners = null, string[] members = null, Stream groupLogo = null,
            bool isPrivate = false, bool createTeam = false, int retryCount = 10, int delay = 500, AzureEnvironment azureEnvironment = AzureEnvironment.Production,
            Enums.Office365Geography? preferredDataLocation = null, Guid[] assignedLabels = null, bool welcomeEmailDisabled = false)
        {
            UnifiedGroupEntity result = null;

            if (string.IsNullOrEmpty(displayName))
            {
                throw new ArgumentNullException(nameof(displayName));
            }

            if (string.IsNullOrEmpty(mailNickname))
            {
                throw new ArgumentNullException(nameof(mailNickname));
            }

            if (string.IsNullOrEmpty(accessToken))
            {
                throw new ArgumentNullException(nameof(accessToken));
            }

            var labels = new List<AssignedLabel>();
            if(assignedLabels != null)
            {
                foreach (var label in assignedLabels)
                {
                    if (!Guid.Empty.Equals(label))
                    {
                        labels.Add(new AssignedLabel
                        {
                            LabelId = label.ToString()
                        });
                    }
                }
            }
            

            try
            {
                // Use a synchronous model to invoke the asynchronous process
                result = Task.Run(async () =>
                {
                    var group = new UnifiedGroupEntity();

                    var graphClient = CreateGraphClient(accessToken, retryCount, delay, azureEnvironment);

                    // Prepare the group resource object
                    var newGroup = new GroupExtended
                    {
                        DisplayName = displayName,
                        Description = string.IsNullOrEmpty(description) ? null : description,
                        MailNickname = mailNickname,
                        MailEnabled = true,
                        SecurityEnabled = false,
                        Visibility = isPrivate == true ? "Private" : "Public",
                        GroupTypes = new List<string> { "Unified" }
                    };

                    if (labels.Any())
                    {
                        newGroup.AssignedLabels = labels;
                    }

                    if (preferredDataLocation.HasValue)
                    {
                        newGroup.PreferredDataLocation = preferredDataLocation.Value.ToString();
                    }

                    if (owners != null && owners.Length > 0)
                    {
                        var users = GetUsers(graphClient, owners);
                        if (users != null && users.Count > 0)
                        {
                            newGroup.OwnersODataBind = users.Select(u => $"https://{AuthenticationManager.GetGraphEndPoint(azureEnvironment)}/v1.0/users/{u.Id}").ToArray();
                        }
                    }

                    if (members != null && members.Length > 0)
                    {
                        var users = GetUsers(graphClient, members);
                        if (users != null && users.Count > 0)
                        {
                            newGroup.MembersODataBind = users.Select(u => $"https://{AuthenticationManager.GetGraphEndPoint(azureEnvironment)}/v1.0/users/{u.Id}").ToArray();
                        }
                    }

                    if (welcomeEmailDisabled)
                    {
                        if (newGroup.AdditionalData == null)
                        {
                            newGroup.AdditionalData = new Dictionary<string, object>();
                        }
                        newGroup.AdditionalData.Add("resourceBehaviorOptions", new string[] { "WelcomeEmailDisabled" });
                    }

                    Microsoft.Graph.Group addedGroup = null;
                    string modernSiteUrl = null;

                    // Add the group to the collection of groups (if it does not exist)
                    if (addedGroup == null)
                    {
                        addedGroup = await graphClient.Groups.Request().AddAsync(newGroup);

                        if (addedGroup != null)
                        {
                            group.DisplayName = addedGroup.DisplayName;
                            group.Description = addedGroup.Description;
                            group.GroupId = addedGroup.Id;
                            group.Mail = addedGroup.Mail;
                            group.MailNickname = addedGroup.MailNickname;

                            int imageRetryCount = retryCount;

                            if (groupLogo != null)
                            {
                                using (var memGroupLogo = new MemoryStream())
                                {
                                    groupLogo.CopyTo(memGroupLogo);

                                    while (imageRetryCount > 0)
                                    {
                                        bool groupLogoUpdated = false;
                                        memGroupLogo.Position = 0;

                                        using (var tempGroupLogo = new MemoryStream())
                                        {
                                            memGroupLogo.CopyTo(tempGroupLogo);
                                            tempGroupLogo.Position = 0;

                                            try
                                            {
                                                groupLogoUpdated = UpdateUnifiedGroup(addedGroup.Id, accessToken, groupLogo: tempGroupLogo, azureEnvironment: azureEnvironment);
                                            }
                                            catch
                                            {
                                                // Skip any exception and simply retry
                                            }
                                        }

                                        // In case of failure retry up to 10 times, with 500ms delay in between
                                        if (!groupLogoUpdated)
                                        {
                                            // Pop up the delay for the group image
                                            await Task.Delay(delay * (retryCount - imageRetryCount));
                                            imageRetryCount--;
                                        }
                                        else
                                        {
                                            break;
                                        }
                                    }
                                }
                            }

                            int driveRetryCount = retryCount;

                            while (driveRetryCount > 0 && string.IsNullOrEmpty(modernSiteUrl))
                            {
                                try
                                {
                                    modernSiteUrl = GetUnifiedGroupSiteUrl(addedGroup.Id, accessToken, azureEnvironment: azureEnvironment);
                                }
                                catch
                                {
                                    // Skip any exception and simply retry
                                }

                                // In case of failure retry up to 10 times, with 500ms delay in between
                                if (string.IsNullOrEmpty(modernSiteUrl))
                                {
                                    await Task.Delay(delay * (retryCount - driveRetryCount));
                                    driveRetryCount--;
                                }
                            }

                            group.SiteUrl = modernSiteUrl;
                        }
                    }

                    if (createTeam)
                    {
                        await CreateTeam(group.GroupId, accessToken, azureEnvironment: azureEnvironment);
                    }

                    return (group);

                }).GetAwaiter().GetResult();
            }
            catch (ServiceException ex)
            {
                Log.Error(Constants.LOGGING_SOURCE, CoreResources.GraphExtensions_ErrorOccured, ex.Error.Message);
                throw;
            }
            return (result);
        }

        /// <summary>
        /// Updates the members of a Microsoft 365 Group
        /// </summary>
        /// <param name="members">UPNs of users that need to be added as a member to the group</param>
        /// <param name="graphClient">GraphClient instance to use to communicate with the Microsoft Graph</param>
        /// <param name="groupId">Id of the group which needs the owners added</param>
        /// <param name="removeOtherMembers">If set to true, all existing members which are not specified through <paramref name="members"/> will be removed as a member from the group</param>
        private static async Task UpdateMembers(string[] members, GraphServiceClient graphClient, string groupId, bool removeOtherMembers)
        {
            if (members != null && members.Length > 0)
            {
                foreach (var m in members)
                {
                    // Search for the user object
                    var memberQuery = await graphClient.Users
                        .Request()
                        .Filter($"userPrincipalName eq '{Uri.EscapeDataString(m.Replace("'", "''"))}'")
                        .GetAsync();

                    var member = memberQuery.FirstOrDefault();

                    if (member != null)
                    {
                        try
                        {
                            // And if any, add it to the collection of group's owners
                            await graphClient.Groups[groupId].Members.References.Request().AddAsync(member);
                        }
                        catch (Exception ex)
                        {
                            if (ex.Message.Contains("Request_BadRequest") &&
                                ex.Message.Contains("added object references already exist"))
                            {
                                // Skip any already existing member
                            }
                            else
                            {
#pragma warning disable CA2200
                                throw ex;
#pragma warning restore CA2200
                            }
                        }
                    }
                }
            }
            // Check if all other members not provided should be removed
            if (!removeOtherMembers)
            {
                return;
            }

            // Remove any leftover member
            var fullListOfMembers = await graphClient.Groups[groupId].Members.Request().Select("userPrincipalName, Id").GetAsync();
            var pageExists = true;

            while (pageExists)
            {
                foreach (var member in fullListOfMembers)
                {
                    var currentMemberPrincipalName = (member as Microsoft.Graph.User)?.UserPrincipalName;
                    if (!string.IsNullOrEmpty(currentMemberPrincipalName) &&
                        !members.Contains(currentMemberPrincipalName, StringComparer.InvariantCultureIgnoreCase))
                    {
                        try
                        {
                            // If it is not in the list of current owners, just remove it
                            await graphClient.Groups[groupId].Members[member.Id].Reference.Request().DeleteAsync();
                        }
                        catch (ServiceException ex)
                        {
                            if (ex.Error.Code == "Request_BadRequest")
                            {
                                // Skip any failing removal
                            }
                            else
                            {
                                throw ex;
                            }
                        }
                    }
                }

                if (fullListOfMembers.NextPageRequest != null)
                {
                    fullListOfMembers = await fullListOfMembers.NextPageRequest.GetAsync();
                }
                else
                {
                    pageExists = false;
                }
            }
        }

        /// <summary>
        /// Updates the owners of a Microsoft 365 Group
        /// </summary>
        /// <param name="owners">UPNs of users that need to be added as a owner to the group</param>
        /// <param name="graphClient">GraphClient instance to use to communicate with the Microsoft Graph</param>
        /// <param name="groupId">Id of the group which needs the owners added</param>
        /// <param name="removeOtherOwners">If set to true, all existing owners which are not specified through <paramref name="owners"/> will be removed as an owner from the group</param>
        private static async Task UpdateOwners(string[] owners, GraphServiceClient graphClient, string groupId, bool removeOtherOwners)
        {
            foreach (var o in owners)
            {
                // Search for the user object
                var ownerQuery = await graphClient.Users
                    .Request()
                    .Filter($"userPrincipalName eq '{Uri.EscapeDataString(o.Replace("'", "''"))}'")
                    .GetAsync();

                var owner = ownerQuery.FirstOrDefault();

                if (owner != null)
                {
                    try
                    {
                        // And if any, add it to the collection of group's owners
                        await graphClient.Groups[groupId].Owners.References.Request().AddAsync(owner);
                    }
                    catch (Exception ex)
                    {
                        if (ex.Message.Contains("Request_BadRequest") &&
                            ex.Message.Contains("added object references already exist"))
                        {
                            // Skip any already existing owner
                        }
                        else
                        {
#pragma warning disable CA2200
                            throw ex;
#pragma warning restore CA2200
                        }
                    }
                }
            }

            // Check if all owners which have not been provided should be removed
            if (!removeOtherOwners)
            {
                return;
            }

            // Remove any leftover owner
            var fullListOfOwners = await graphClient.Groups[groupId].Owners.Request().Select("userPrincipalName, Id").GetAsync();
            var pageExists = true;

            while (pageExists)
            {
                foreach (var owner in fullListOfOwners)
                {
                    var currentOwnerPrincipalName = (owner as Microsoft.Graph.User)?.UserPrincipalName;
                    if (!string.IsNullOrEmpty(currentOwnerPrincipalName) &&
                        !owners.Contains(currentOwnerPrincipalName, StringComparer.InvariantCultureIgnoreCase))
                    {
                        try
                        {
                            // If it is not in the list of current owners, just remove it
                            await graphClient.Groups[groupId].Owners[owner.Id].Reference.Request().DeleteAsync();
                        }
                        catch (ServiceException ex)
                        {
                            if (ex.Error.Code == "Request_BadRequest")
                            {
                                // Skip any failing removal
                            }
                            else
                            {
                                throw ex;
                            }
                        }
                    }
                }

                if (fullListOfOwners.NextPageRequest != null)
                {
                    fullListOfOwners = await fullListOfOwners.NextPageRequest.GetAsync();
                }
                else
                {
                    pageExists = false;
                }
            }
        }

        /// <summary>
        /// Sets the visibility of a Group
        /// </summary>
        /// <param name="groupId">Id of the Microsoft 365 Group to set the visibility state for</param>
        /// <param name="accessToken">The OAuth 2.0 Access Token to use for invoking the Microsoft Graph</param>
        /// <param name="hideFromAddressLists">True if the group should not be displayed in certain parts of the Outlook UI: the Address Book, address lists for selecting message recipients, and the Browse Groups dialog for searching groups; otherwise, false. Default value is false.</param>
        /// <param name="hideFromOutlookClients">True if the group should not be displayed in Outlook clients, such as Outlook for Windows and Outlook on the web; otherwise, false. Default value is false.</param>
        /// <param name="azureEnvironment">Defines the Azure Cloud Deployment. This is used to determine the MS Graph EndPoint to call which differs per Azure Cloud deployments. Defaults to Production (graph.microsoft.com).</param>
        public static void SetUnifiedGroupVisibility(string groupId, string accessToken, bool? hideFromAddressLists, bool? hideFromOutlookClients, AzureEnvironment azureEnvironment = AzureEnvironment.Production)
        {
            if (String.IsNullOrEmpty(groupId))
            {
                throw new ArgumentNullException(nameof(groupId));
            }
            if (String.IsNullOrEmpty(accessToken))
            {
                throw new ArgumentNullException(nameof(accessToken));
            }

            // Ensure there's something to update
            if (!hideFromAddressLists.HasValue && !hideFromOutlookClients.HasValue)
            {
                return;
            }

            try
            {
                // PATCH https://graph.microsoft.com/v1.0/groups/{id}
                string updateGroupUrl = $"{GraphHttpClient.GetGraphEndPointUrl(azureEnvironment)}groups/{groupId}";
                var groupRequest = new Model.Group
                {
                    HideFromAddressLists = hideFromAddressLists,
                    HideFromOutlookClients = hideFromOutlookClients
                };

                var response = GraphHttpClient.MakePatchRequestForString(
                    requestUrl: updateGroupUrl,
                    content: JsonConvert.SerializeObject(groupRequest),
                    contentType: "application/json",
                    accessToken: accessToken);
            }
            catch (ServiceException ex)
            {
                Log.Error(Constants.LOGGING_SOURCE, CoreResources.GraphExtensions_ErrorOccured, ex.Error.Message);
                throw;
            }
        }

        /// <summary>
        /// Renews the Office 365 Group by extending its expiration with the number of days defined in the group expiration policy set on the Azure Active Directory
        /// </summary>
        /// <param name="groupId">The ID of the Office 365 Group</param>
        /// <param name="accessToken">The OAuth 2.0 Access Token to use for invoking the Microsoft Graph</param>
        /// <param name="retryCount">Number of times to retry the request in case of throttling</param>
        /// <param name="delay">Milliseconds to wait before retrying the request. The delay will be increased (doubled) every retry.</param>
        /// <param name="azureEnvironment">Defines the Azure Cloud Deployment. This is used to determine the MS Graph EndPoint to call which differs per Azure Cloud deployments. Defaults to Production (graph.microsoft.com).</param>
        public static void RenewUnifiedGroup(string groupId,
                                             string accessToken, int retryCount = 10, int delay = 500, AzureEnvironment azureEnvironment = AzureEnvironment.Production)
        {
            try
            {
                // Use a synchronous model to invoke the asynchronous process
                Task.Run(async () =>
                {
                    var graphClient = CreateGraphClient(accessToken, retryCount, delay, azureEnvironment);

                    await graphClient.Groups[groupId]
                        .Renew()
                        .Request()
                        .PostAsync();
                }).GetAwaiter().GetResult();
            }
            catch (ServiceException ex)
            {
                Log.Error(Constants.LOGGING_SOURCE, CoreResources.GraphExtensions_ErrorOccured, ex.Error.Message);
                throw;
            }
        }

        /// <summary>
        /// Updates the logo, members or visibility state of an Office 365 Group
        /// </summary>
        /// <param name="groupId">The ID of the Office 365 Group</param>
        /// <param name="displayName">The Display Name for the Office 365 Group</param>
        /// <param name="description">The Description for the Office 365 Group</param>
        /// <param name="owners">A list of UPNs for group owners, if any, to be added to the site</param>
        /// <param name="members">A list of UPNs for group members, if any, to be added to the site</param>
        /// <param name="isPrivate">Defines whether the group will be private or public, optional with default false (i.e. public)</param>
        /// <param name="createTeam">Defines whether to create MS Teams team associated with the group</param>
        /// <param name="groupLogo">The binary stream of the logo for the Office 365 Group</param>
        /// <param name="accessToken">The OAuth 2.0 Access Token to use for invoking the Microsoft Graph</param>
        /// <param name="retryCount">Number of times to retry the request in case of throttling</param>
        /// <param name="delay">Milliseconds to wait before retrying the request. The delay will be increased (doubled) every retry</param>
        /// <param name="azureEnvironment">Defines the Azure Cloud Deployment. This is used to determine the MS Graph EndPoint to call which differs per Azure Cloud deployments. Defaults to Production (graph.microsoft.com).</param>
        /// <returns>Declares whether the Office 365 Group has been updated or not</returns>
        public static bool UpdateUnifiedGroup(string groupId,
            string accessToken, int retryCount = 10, int delay = 500,
            string displayName = null, string description = null, string[] owners = null, string[] members = null,
            Stream groupLogo = null, bool? isPrivate = null, bool createTeam = false, AzureEnvironment azureEnvironment = AzureEnvironment.Production)
        {
            bool result;
            try
            {
                // Use a synchronous model to invoke the asynchronous process
                result = Task.Run(async () =>
                {
                    var graphClient = CreateGraphClient(accessToken, retryCount, delay, azureEnvironment);

                    var groupToUpdate = await graphClient.Groups[groupId]
                        .Request()
                        .GetAsync();

                    // Workaround for the PATCH request, needed after update to Graph Library
                    var clonedGroup = new Group
                    {
                        Id = groupToUpdate.Id
                    };

                    #region Logic to update the group DisplayName and Description

                    var updateGroup = false;
                    var groupUpdated = false;

                    // Check if we have to update the DisplayName
                    if (!String.IsNullOrEmpty(displayName) && groupToUpdate.DisplayName != displayName)
                    {
                        clonedGroup.DisplayName = displayName;
                        updateGroup = true;
                    }

                    // Check if we have to update the Description
                    if (!String.IsNullOrEmpty(description) && groupToUpdate.Description != description)
                    {
                        clonedGroup.Description = description;
                        updateGroup = true;
                    }

                    // Check if visibility has changed for the Group
                    bool existingIsPrivate = groupToUpdate.Visibility == "Private";
                    if (isPrivate.HasValue && existingIsPrivate != isPrivate)
                    {
                        clonedGroup.Visibility = isPrivate == true ? "Private" : "Public";
                        updateGroup = true;
                    }

                    // Check if we need to update owners
                    if (owners != null && owners.Length > 0)
                    {
                        // For each and every owner
                        await UpdateOwners(owners, graphClient, groupToUpdate.Id, true);
                        updateGroup = true;
                    }

                    // Check if we need to update members
                    if (members != null && members.Length > 0)
                    {
                        // For each and every owner
                        await UpdateMembers(members, graphClient, groupToUpdate.Id, true);
                        updateGroup = true;
                    }

                    if (createTeam)
                    {
                        await CreateTeam(groupId, accessToken);
                        updateGroup = true;
                    }

                    // If the Group has to be updated, just do it
                    if (updateGroup)
                    {
                        var updatedGroup = await graphClient.Groups[groupId]
                            .Request()
                            .UpdateAsync(clonedGroup);

                        groupUpdated = true;
                    }

                    #endregion

                    #region Logic to update the group Logo

                    var logoUpdated = false;

                    if (groupLogo != null)
                    {
                        await graphClient.Groups[groupId].Photo.Content.Request().PutAsync(groupLogo);
                        logoUpdated = true;
                    }

                    #endregion

                    // If any of the previous update actions has been completed
                    return (groupUpdated || logoUpdated);

                }).GetAwaiter().GetResult();
            }
            catch (ServiceException ex)
            {
                Log.Error(Constants.LOGGING_SOURCE, CoreResources.GraphExtensions_ErrorOccured, ex.Error.Message);
                throw;
            }
            return (result);
        }

        /// <summary>
        /// Creates a new Office 365 Group (i.e. Unified Group) with its backing Modern SharePoint Site
        /// </summary>
        /// <param name="displayName">The Display Name for the Office 365 Group</param>
        /// <param name="description">The Description for the Office 365 Group</param>
        /// <param name="mailNickname">The Mail Nickname for the Office 365 Group</param>
        /// <param name="accessToken">The OAuth 2.0 Access Token to use for invoking the Microsoft Graph</param>
        /// <param name="owners">A list of UPNs for group owners, if any</param>
        /// <param name="members">A list of UPNs for group members, if any</param>
        /// <param name="groupLogoPath">The path of the logo for the Office 365 Group</param>
        /// <param name="isPrivate">Defines whether the group will be private or public, optional with default false (i.e. public)</param>
        /// <param name="createTeam">Defines whether to create MS Teams team associated with the group</param>
        /// <param name="retryCount">Number of times to retry the request in case of throttling</param>
        /// <param name="delay">Milliseconds to wait before retrying the request. The delay will be increased (doubled) every retry</param>
        /// <param name="azureEnvironment">Defines the Azure Cloud Deployment. This is used to determine the MS Graph EndPoint to call which differs per Azure Cloud deployments. Defaults to Production (graph.microsoft.com).</param>
        /// <param name="preferredDataLocation">Defines the codes of geographies in which there is Office 365 presence. Used for multi-geo enabled tenants. List with available geographies is available at https://docs.microsoft.com/office365/enterprise/multi-geo-add-group-with-pdl#geo-location-codes.</param>
        /// <returns>The just created Office 365 Group</returns>
        public static UnifiedGroupEntity CreateUnifiedGroup(string displayName, string description, string mailNickname,
            string accessToken, string[] owners = null, string[] members = null, string groupLogoPath = null,
            bool isPrivate = false, bool createTeam = false, int retryCount = 10, int delay = 500, AzureEnvironment azureEnvironment = AzureEnvironment.Production, Enums.Office365Geography? preferredDataLocation = null)
        {
            if (!string.IsNullOrEmpty(groupLogoPath) && !System.IO.File.Exists(groupLogoPath))
            {
                throw new FileNotFoundException(CoreResources.GraphExtensions_GroupLogoFileDoesNotExist, groupLogoPath);
            }
            else if (!string.IsNullOrEmpty(groupLogoPath))
            {
                using (var groupLogoStream = new FileStream(groupLogoPath, FileMode.Open, FileAccess.Read, FileShare.Read))
                {
                    return (CreateUnifiedGroup(displayName, description,
                        mailNickname, accessToken, owners, members,
                        groupLogo: groupLogoStream, isPrivate: isPrivate,
                        createTeam: createTeam, retryCount: retryCount, delay: delay, azureEnvironment: azureEnvironment, preferredDataLocation: preferredDataLocation));
                }
            }
            else
            {
                return (CreateUnifiedGroup(displayName, description,
                    mailNickname, accessToken, owners, members,
                    groupLogo: null, isPrivate: isPrivate,
                    createTeam: createTeam, retryCount: retryCount, delay: delay, azureEnvironment: azureEnvironment, preferredDataLocation: preferredDataLocation));
            }
        }

        /// <summary>
        /// Creates a new Office 365 Group (i.e. Unified Group) with its backing Modern SharePoint Site
        /// </summary>
        /// <param name="displayName">The Display Name for the Office 365 Group</param>
        /// <param name="description">The Description for the Office 365 Group</param>
        /// <param name="mailNickname">The Mail Nickname for the Office 365 Group</param>
        /// <param name="accessToken">The OAuth 2.0 Access Token to use for invoking the Microsoft Graph</param>
        /// <param name="owners">A list of UPNs for group owners, if any</param>
        /// <param name="members">A list of UPNs for group members, if any</param>
        /// <param name="isPrivate">Defines whether the group will be private or public, optional with default false (i.e. public)</param>
        /// <param name="createTeam">Defines whether to create MS Teams team associated with the group</param>
        /// <param name="retryCount">Number of times to retry the request in case of throttling</param>
        /// <param name="delay">Milliseconds to wait before retrying the request. The delay will be increased (doubled) every retry</param>
        /// <param name="azureEnvironment">Defines the Azure Cloud Deployment. This is used to determine the MS Graph EndPoint to call which differs per Azure Cloud deployments. Defaults to Production (graph.microsoft.com).</param>
        /// <param name="preferredDataLocation">Defines the codes of geographies in which there is Office 365 presence. Used for multi-geo enabled tenants. List with available geographies is available at https://docs.microsoft.com/office365/enterprise/multi-geo-add-group-with-pdl#geo-location-codes.</param>
        /// <returns>The just created Office 365 Group</returns>
        public static UnifiedGroupEntity CreateUnifiedGroup(string displayName, string description, string mailNickname,
            string accessToken, string[] owners = null, string[] members = null,
            bool isPrivate = false, bool createTeam = false, int retryCount = 10, int delay = 500, AzureEnvironment azureEnvironment = AzureEnvironment.Production, Enums.Office365Geography? preferredDataLocation = null)
        {
            return (CreateUnifiedGroup(displayName, description,
                mailNickname, accessToken, owners, members,
                groupLogo: null, isPrivate: isPrivate,
                createTeam: createTeam, retryCount: retryCount, delay: delay, azureEnvironment: azureEnvironment, preferredDataLocation: preferredDataLocation));
        }

        /// <summary>
        /// Deletes an Office 365 Group (i.e. Unified Group)
        /// </summary>
        /// <param name="groupId">The ID of the Office 365 Group</param>
        /// <param name="accessToken">The OAuth 2.0 Access Token to use for invoking the Microsoft Graph</param>
        /// <param name="retryCount">Number of times to retry the request in case of throttling</param>
        /// <param name="delay">Milliseconds to wait before retrying the request. The delay will be increased (doubled) every retry</param>
        /// <param name="azureEnvironment">Defines the Azure Cloud Deployment. This is used to determine the MS Graph EndPoint to call which differs per Azure Cloud deployments. Defaults to Production (graph.microsoft.com).</param>
        public static void DeleteUnifiedGroup(string groupId, string accessToken, int retryCount = 10, int delay = 500, AzureEnvironment azureEnvironment = AzureEnvironment.Production)
        {
            if (string.IsNullOrEmpty(groupId))
            {
                throw new ArgumentNullException(nameof(groupId));
            }

            if (string.IsNullOrEmpty(accessToken))
            {
                throw new ArgumentNullException(nameof(accessToken));
            }
            try
            {
                // Use a synchronous model to invoke the asynchronous process
                Task.Run(async () =>
                {
                    var graphClient = CreateGraphClient(accessToken, retryCount, delay, azureEnvironment);
                    await graphClient.Groups[groupId].Request().DeleteAsync();

                }).GetAwaiter().GetResult();
            }
            catch (ServiceException ex)
            {
                Log.Error(Constants.LOGGING_SOURCE, CoreResources.GraphExtensions_ErrorOccured, ex.Error.Message);
                throw;
            }
        }

        /// <summary>
        /// Get an Office 365 Group (i.e. Unified Group) by Id
        /// </summary>
        /// <param name="groupId">The ID of the Office 365 Group</param>
        /// <param name="accessToken">The OAuth 2.0 Access Token to use for invoking the Microsoft Graph</param>
        /// <param name="includeSite">Defines whether to return details about the Modern SharePoint Site backing the group. Default is true.</param>
        /// <param name="retryCount">Number of times to retry the request in case of throttling</param>
        /// <param name="delay">Milliseconds to wait before retrying the request. The delay will be increased (doubled) every retry</param>
        /// <param name="includeClassification">Defines whether to return classification value of the unified group. Default is false.</param>
        /// <param name="includeHasTeam">Defines whether to check for each unified group if it has a Microsoft Team provisioned for it. Default is false.</param>
        /// <param name="azureEnvironment">Defines the Azure Cloud Deployment. This is used to determine the MS Graph EndPoint to call which differs per Azure Cloud deployments. Defaults to Production (graph.microsoft.com).</param>
        public static UnifiedGroupEntity GetUnifiedGroup(string groupId, string accessToken, int retryCount = 10, int delay = 500, bool includeSite = true, bool includeClassification = false, bool includeHasTeam = false, AzureEnvironment azureEnvironment = AzureEnvironment.Production)
        {
            if (string.IsNullOrEmpty(groupId))
            {
                throw new ArgumentNullException(nameof(groupId));
            }

            if (string.IsNullOrEmpty(accessToken))
            {
                throw new ArgumentNullException(nameof(accessToken));
            }

            UnifiedGroupEntity result = null;
            try
            {
                // Use a synchronous model to invoke the asynchronous process
                result = Task.Run(async () =>
                {
                    UnifiedGroupEntity group = null;

                    var graphClient = CreateGraphClient(accessToken, retryCount, delay, azureEnvironment);

                    var g = await graphClient.Groups[groupId].Request().GetAsync();

                    if (g.GroupTypes.Contains("Unified"))
                    {
                        group = new UnifiedGroupEntity
                        {
                            GroupId = g.Id,
                            DisplayName = g.DisplayName,
                            Description = g.Description,
                            Mail = g.Mail,
                            MailNickname = g.MailNickname,
                            Visibility = g.Visibility
                        };
                        if (includeSite)
                        {
                            try
                            {
                                group.SiteUrl = GetUnifiedGroupSiteUrl(groupId, accessToken);
                            }
                            catch (ServiceException e)
                            {
                                group.SiteUrl = e.Error.Message;
                            }
                        }

                        if (includeClassification)
                        {
                            group.Classification = g.Classification;
                        }

                        if (includeHasTeam)
                        {
                            group.HasTeam = HasTeamsTeam(group.GroupId, accessToken, azureEnvironment);
                        }
                    }

                    return (group);

                }).GetAwaiter().GetResult();
            }
            catch (ServiceException ex)
            {
                Log.Error(Constants.LOGGING_SOURCE, CoreResources.GraphExtensions_ErrorOccured, ex.Error.Message);
                throw;
            }
            return (result);
        }

        /// <summary>
        /// Returns all the Office 365 Groups in the current Tenant based on a startIndex. IncludeSite adds additional properties about the Modern SharePoint Site backing the group
        /// </summary>
        /// <param name="accessToken">The OAuth 2.0 Access Token to use for invoking the Microsoft Graph</param>
        /// <param name="displayName">The DisplayName of the Office 365 Group</param>
        /// <param name="mailNickname">The MailNickname of the Office 365 Group</param>
        /// <param name="startIndex">Not relevant anymore</param>
        /// <param name="endIndex">Not relevant anymore</param>
        /// <param name="includeSite">Defines whether to return details about the Modern SharePoint Site backing the group. Default is true.</param>
        /// <param name="retryCount">Number of times to retry the request in case of throttling</param>
        /// <param name="delay">Milliseconds to wait before retrying the request. The delay will be increased (doubled) every retry</param>
        /// <param name="includeClassification">Defines whether or not to return details about the Modern Site classification value.</param>
        /// <param name="includeHasTeam">Defines whether to check for each unified group if it has a Microsoft Team provisioned for it. Default is false.</param>
        /// <param name="azureEnvironment">Defines the Azure Cloud Deployment. This is used to determine the MS Graph EndPoint to call which differs per Azure Cloud deployments. Defaults to Production (graph.microsoft.com).</param>
        /// <returns>An IList of SiteEntity objects</returns>
        [Obsolete("ListUnifiedGroups is deprecated, please use GetUnifiedGroups instead.")]
        public static List<UnifiedGroupEntity> ListUnifiedGroups(string accessToken,
            string displayName = null, string mailNickname = null,
            int startIndex = 0, int endIndex = 999, bool includeSite = true,
            int retryCount = 10, int delay = 500, bool includeClassification = false,
            bool includeHasTeam = false, AzureEnvironment azureEnvironment = AzureEnvironment.Production)
        {
            if (string.IsNullOrEmpty(accessToken))
            {
                throw new ArgumentNullException(nameof(accessToken));
            }

            List<UnifiedGroupEntity> result = null;
            try
            {
                // Use a synchronous model to invoke the asynchronous process
                result = Task.Run(async () =>
                {
                    List<UnifiedGroupEntity> groups = new List<UnifiedGroupEntity>();

                    var graphClient = CreateGraphClient(accessToken, retryCount, delay, azureEnvironment);

                    // Apply the DisplayName filter, if any
                    var displayNameFilter = !string.IsNullOrEmpty(displayName) ? $" and (DisplayName eq '{Uri.EscapeDataString(displayName.Replace("'", "''"))}')" : string.Empty;
                    var mailNicknameFilter = !string.IsNullOrEmpty(mailNickname) ? $" and (MailNickname eq '{Uri.EscapeDataString(mailNickname.Replace("'", "''"))}')" : string.Empty;

                    var pagedGroups = await graphClient.Groups
                        .Request()
                        .Filter($"groupTypes/any(grp: grp eq 'Unified'){displayNameFilter}{mailNicknameFilter}")
                        .Top(endIndex)
                        .GetAsync();

                    Int32 pageCount = 0;
                    Int32 currentIndex = 0;

                    while (true)
                    {
                        pageCount++;

                        foreach (var g in pagedGroups)
                        {
                            currentIndex++;

                            if (currentIndex >= startIndex)
                            {
                                var group = new UnifiedGroupEntity
                                {
                                    GroupId = g.Id,
                                    DisplayName = g.DisplayName,
                                    Description = g.Description,
                                    Mail = g.Mail,
                                    MailNickname = g.MailNickname,
                                    Visibility = g.Visibility
                                };

                                if (includeSite)
                                {
                                    try
                                    {
                                        group.SiteUrl = GetUnifiedGroupSiteUrl(g.Id, accessToken);
                                    }
                                    catch (ServiceException e)
                                    {
                                        group.SiteUrl = e.Error.Message;
                                    }
                                }

                                if (includeClassification)
                                {
                                    group.Classification = g.Classification;
                                }

                                if (includeHasTeam)
                                {
                                    group.HasTeam = HasTeamsTeam(group.GroupId, accessToken);
                                }

                                groups.Add(group);
                            }
                        }

                        if (pagedGroups.NextPageRequest != null && groups.Count < endIndex)
                        {
                            pagedGroups = await pagedGroups.NextPageRequest.GetAsync();
                        }
                        else
                        {
                            break;
                        }
                    }

                    return (groups);
                }).GetAwaiter().GetResult();
            }
            catch (ServiceException ex)
            {
                Log.Error(Constants.LOGGING_SOURCE, CoreResources.GraphExtensions_ErrorOccured, ex.Error.Message);
                throw;
            }
            return (result);
        }

        /// <summary>
        /// Returns all the Office 365 Groups in the current Tenant based on a startIndex. IncludeSite adds additional properties about the Modern SharePoint Site backing the group
        /// </summary>
        /// <param name="accessToken">The OAuth 2.0 Access Token to use for invoking the Microsoft Graph</param>
        /// <param name="displayName">The DisplayName of the Office 365 Group</param>
        /// <param name="mailNickname">The MailNickname of the Office 365 Group</param>
        /// <param name="startIndex">If not specified, method will start with the first group.</param>
        /// <param name="endIndex">If not specified, method will return all groups.</param>
        /// <param name="includeSite">Defines whether to return details about the Modern SharePoint Site backing the group. Default is true.</param>
        /// <param name="retryCount">Number of times to retry the request in case of throttling</param>
        /// <param name="delay">Milliseconds to wait before retrying the request. The delay will be increased (doubled) every retry</param>
        /// <param name="includeClassification">Defines whether or not to return details about the Modern Site classification value.</param>
        /// <param name="includeHasTeam">Defines whether to check for each unified group if it has a Microsoft Team provisioned for it. Default is false.</param>
        /// <param name="pageSize">Page size used for the individual requests to Micrsoft Graph. Defaults to 999 which is currently the maximum value.</param>
        /// <param name="azureEnvironment">Defines the Azure Cloud Deployment. This is used to determine the MS Graph EndPoint to call which differs per Azure Cloud deployments. Defaults to Production (graph.microsoft.com).</param>
        /// <returns>An IList of SiteEntity objects</returns>
        public static List<UnifiedGroupEntity> GetUnifiedGroups(string accessToken,
            string displayName = null, string mailNickname = null,
            int startIndex = 0, int? endIndex = null, bool includeSite = true,
            int retryCount = 10, int delay = 500, bool includeClassification = false, bool includeHasTeam = false, int pageSize = 999, AzureEnvironment azureEnvironment = AzureEnvironment.Production)
        {
            if (string.IsNullOrEmpty(accessToken))
            {
                throw new ArgumentNullException(nameof(accessToken));
            }

            List<UnifiedGroupEntity> result = null;
            try
            {
                // Use a synchronous model to invoke the asynchronous process
                result = Task.Run(async () =>
                {
                    List<UnifiedGroupEntity> groups = new List<UnifiedGroupEntity>();

                    var graphClient = CreateGraphClient(accessToken, retryCount, delay, azureEnvironment);

                    // Apply the DisplayName filter, if any
                    var displayNameFilter = !string.IsNullOrEmpty(displayName) ? $" and (DisplayName eq '{Uri.EscapeDataString(displayName.Replace("'", "''"))}')" : string.Empty;
                    var mailNicknameFilter = !string.IsNullOrEmpty(mailNickname) ? $" and (MailNickname eq '{Uri.EscapeDataString(mailNickname.Replace("'", "''"))}')" : string.Empty;

                    var pagedGroups = await graphClient.Groups
                        .Request()
                        .Filter($"groupTypes/any(grp: grp eq 'Unified'){displayNameFilter}{mailNicknameFilter}")
                        .Top(pageSize)
                        .GetAsync();

                    Int32 pageCount = 0;
                    Int32 currentIndex = 0;

                    while (true)
                    {
                        pageCount++;

                        foreach (var g in pagedGroups)
                        {
                            currentIndex++;

                            if (currentIndex >= startIndex)
                            {
                                var group = new UnifiedGroupEntity
                                {
                                    GroupId = g.Id,
                                    DisplayName = g.DisplayName,
                                    Description = g.Description,
                                    Mail = g.Mail,
                                    MailNickname = g.MailNickname,
                                    Visibility = g.Visibility
                                };

                                if (includeSite)
                                {
                                    try
                                    {
                                        group.SiteUrl = GetUnifiedGroupSiteUrl(g.Id, accessToken);
                                    }
                                    catch (ServiceException e)
                                    {
                                        group.SiteUrl = e.Error.Message;
                                    }
                                }

                                if (includeClassification)
                                {
                                    group.Classification = g.Classification;
                                }

                                if (includeHasTeam)
                                {
                                    group.HasTeam = HasTeamsTeam(group.GroupId, accessToken, azureEnvironment);
                                }

                                groups.Add(group);
                            }
                        }

                        if (pagedGroups.NextPageRequest != null && (endIndex == null || groups.Count < endIndex))
                        {
                            pagedGroups = await pagedGroups.NextPageRequest.GetAsync();
                        }
                        else
                        {
                            break;
                        }
                    }

                    return (groups);
                }).GetAwaiter().GetResult();
            }
            catch (ServiceException ex)
            {
                Log.Error(Constants.LOGGING_SOURCE, CoreResources.GraphExtensions_ErrorOccured, ex.Error.Message);
                throw;
            }
            return (result);
        }

        /// <summary>
        /// Returns all the Members of an Office 365 group.
        /// </summary>
        /// <param name="group">The Office 365 group object of type UnifiedGroupEntity</param>
        /// <param name="accessToken">The OAuth 2.0 Access Token to use for invoking the Microsoft Graph</param>
        /// <param name="retryCount">Number of times to retry the request in case of throttling</param>
        /// <param name="delay">Milliseconds to wait before retrying the request. The delay will be increased (doubled) every retry</param>
        /// <param name="azureEnvironment">Defines the Azure Cloud Deployment. This is used to determine the MS Graph EndPoint to call which differs per Azure Cloud deployments. Defaults to Production (graph.microsoft.com).</param>
        /// <returns>Members of an Office 365 group as a list of UnifiedGroupUser entity</returns>
        public static List<UnifiedGroupUser> GetUnifiedGroupMembers(UnifiedGroupEntity group, string accessToken, int retryCount = 10, int delay = 500, AzureEnvironment azureEnvironment = AzureEnvironment.Production)
        {
            List<UnifiedGroupUser> unifiedGroupUsers = null;
            List<User> unifiedGroupGraphUsers = null;
            IGroupMembersCollectionWithReferencesPage groupUsers = null;

            if (String.IsNullOrEmpty(accessToken))
            {
                throw new ArgumentNullException(nameof(accessToken));
            }
            if (group == null)
            {
                throw new ArgumentNullException(nameof(group));
            }

            try
            {
                var result = Task.Run(async () =>
                {
                    var graphClient = CreateGraphClient(accessToken, retryCount, delay, azureEnvironment);

                    // Get the members of an Office 365 group.
                    groupUsers = await graphClient.Groups[group.GroupId].Members.Request().GetAsync();
                    if (groupUsers.CurrentPage != null && groupUsers.CurrentPage.Count > 0)
                    {
                        unifiedGroupGraphUsers = new List<User>();

                        GenerateGraphUserCollection(groupUsers.CurrentPage, unifiedGroupGraphUsers);
                    }

                    // Retrieve users when the results are paged.
                    while (groupUsers.NextPageRequest != null)
                    {
                        groupUsers = groupUsers.NextPageRequest.GetAsync().GetAwaiter().GetResult();
                        if (groupUsers.CurrentPage != null && groupUsers.CurrentPage.Count > 0)
                        {
                            GenerateGraphUserCollection(groupUsers.CurrentPage, unifiedGroupGraphUsers);
                        }
                    }

                    // Create the collection of type OfficeDevPnP 'UnifiedGroupUser' after all users are retrieved, including paged data.
                    if (unifiedGroupGraphUsers != null && unifiedGroupGraphUsers.Count > 0)
                    {
                        unifiedGroupUsers = new List<UnifiedGroupUser>();
                        foreach (User usr in unifiedGroupGraphUsers)
                        {
                            UnifiedGroupUser groupUser = new UnifiedGroupUser
                            {
                                Id = usr.Id,
                                UserPrincipalName = usr.UserPrincipalName != null ? usr.UserPrincipalName : string.Empty,
                                DisplayName = usr.DisplayName != null ? usr.DisplayName : string.Empty,
                                GivenName = usr.GivenName != null ? usr.GivenName : string.Empty,
                                Surname = usr.Surname != null ? usr.Surname : string.Empty,
                                Email = usr.Mail != null ? usr.Mail : string.Empty,
                                MobilePhone = usr.MobilePhone != null ? usr.DisplayName : string.Empty,
                                PreferredLanguage = usr.PreferredLanguage != null ? usr.PreferredLanguage : string.Empty,
                                JobTitle = usr.JobTitle != null ? usr.DisplayName : string.Empty,
                                BusinessPhones = usr.BusinessPhones != null ? usr.BusinessPhones.ToArray() : null
                            };
                            unifiedGroupUsers.Add(groupUser);
                        }
                    }
                    return unifiedGroupUsers;

                }).GetAwaiter().GetResult();
            }
            catch (ServiceException ex)
            {
                Log.Error(Constants.LOGGING_SOURCE, CoreResources.GraphExtensions_ErrorOccured, ex.Error.Message);
                throw;
            }
            return unifiedGroupUsers;
        }

        /// <summary>
        /// Returns all the Members of an Office 365 group (including nested groups).
        /// </summary>
        /// <param name="group"></param>
        /// <param name="accessToken"></param>
        /// <param name="retryCount"></param>
        /// <param name="delay"></param>
        /// <param name="azureEnvironment">Defines the Azure Cloud Deployment. This is used to determine the MS Graph EndPoint to call which differs per Azure Cloud deployments. Defaults to Production (graph.microsoft.com).</param>
        /// <returns></returns>
        public static List<UnifiedGroupUser> GetNestedUnifiedGroupMembers(UnifiedGroupEntity group, string accessToken, int retryCount = 10, int delay = 500, AzureEnvironment azureEnvironment = AzureEnvironment.Production)
        {
            List<UnifiedGroupUser> unifiedGroupUsers = new List<UnifiedGroupUser>();
            List<User> unifiedGroupGraphUsers = null;
            IGroupMembersCollectionWithReferencesPage groupUsers = null;

            if (String.IsNullOrEmpty(accessToken))
            {
                throw new ArgumentNullException(nameof(accessToken));
            }
            if (group == null)
            {
                throw new ArgumentNullException(nameof(group));
            }

            try
            {
                var result = Task.Run(async () =>
                {
                    var graphClient = CreateGraphClient(accessToken, retryCount, delay, azureEnvironment);

                    // Get the members of an Office 365 group.
                    groupUsers = await graphClient.Groups[group.GroupId].Members.Request().GetAsync();
                    if (groupUsers.CurrentPage != null && groupUsers.CurrentPage.Count > 0)
                    {
                        unifiedGroupGraphUsers = new List<User>();

                        GenerateNestedGraphUserCollection(groupUsers.CurrentPage, unifiedGroupGraphUsers, unifiedGroupUsers, accessToken);
                    }

                    // Retrieve users when the results are paged.
                    while (groupUsers.NextPageRequest != null)
                    {
                        groupUsers = groupUsers.NextPageRequest.GetAsync().GetAwaiter().GetResult();
                        if (groupUsers.CurrentPage != null && groupUsers.CurrentPage.Count > 0)
                        {
                            GenerateNestedGraphUserCollection(groupUsers.CurrentPage, unifiedGroupGraphUsers, unifiedGroupUsers, accessToken);
                        }
                    }

                    // Create the collection of type OfficeDevPnP 'UnifiedGroupUser' after all users are retrieved, including paged data.
                    if (unifiedGroupGraphUsers != null && unifiedGroupGraphUsers.Count > 0)
                    {
                        foreach (User usr in unifiedGroupGraphUsers)
                        {
                            UnifiedGroupUser groupUser = new UnifiedGroupUser
                            {
                                Id = usr.Id,
                                UserPrincipalName = usr.UserPrincipalName != null ? usr.UserPrincipalName : string.Empty,
                                DisplayName = usr.DisplayName != null ? usr.DisplayName : string.Empty,
                                GivenName = usr.GivenName != null ? usr.GivenName : string.Empty,
                                Surname = usr.Surname != null ? usr.Surname : string.Empty,
                                Email = usr.Mail != null ? usr.Mail : string.Empty,
                                MobilePhone = usr.MobilePhone != null ? usr.DisplayName : string.Empty,
                                PreferredLanguage = usr.PreferredLanguage != null ? usr.PreferredLanguage : string.Empty,
                                JobTitle = usr.JobTitle != null ? usr.DisplayName : string.Empty,
                                BusinessPhones = usr.BusinessPhones != null ? usr.BusinessPhones.ToArray() : null
                            };
                            unifiedGroupUsers.Add(groupUser);
                        }
                    }
                    return unifiedGroupUsers;

                }).GetAwaiter().GetResult();
            }
            catch (ServiceException ex)
            {
                Log.Error(Constants.LOGGING_SOURCE, CoreResources.GraphExtensions_ErrorOccured, ex.Error.Message);
                throw;
            }
            return unifiedGroupUsers;
        }

        /// <summary>
        /// Adds owners to a Microsoft 365 group
        /// </summary>
        /// <param name="groupId">Id of the Microsoft 365 group to add the owners to</param>
        /// <param name="owners">String array with the UPNs of the users that need to be added as owners to the group</param>
        /// <param name="accessToken">The OAuth 2.0 Access Token to use for invoking the Microsoft Graph</param>
        /// <param name="removeExistingOwners">If true, all existing owners will be removed and only those provided will become owners. If false, existing owners will remain and the ones provided will be added to the list with existing owners.</param>
        /// <param name="retryCount">Number of times to retry the request in case of throttling</param>
        /// <param name="delay">Milliseconds to wait before retrying the request. The delay will be increased (doubled) every retry</param>
        /// <param name="azureEnvironment">Defines the Azure Cloud Deployment. This is used to determine the MS Graph EndPoint to call which differs per Azure Cloud deployments. Defaults to Production (graph.microsoft.com).</param>
        public static void AddUnifiedGroupOwners(string groupId, string[] owners, string accessToken, bool removeExistingOwners = false, int retryCount = 10, int delay = 500, AzureEnvironment azureEnvironment = AzureEnvironment.Production)
        {
            if (String.IsNullOrEmpty(accessToken))
            {
                throw new ArgumentNullException(nameof(accessToken));
            }

            try
            {
                Task.Run(async () =>
                {
                    var graphClient = CreateGraphClient(accessToken, retryCount, delay, azureEnvironment);

                    await UpdateOwners(owners, graphClient, groupId, removeExistingOwners);

                }).GetAwaiter().GetResult();
            }
            catch (ServiceException ex)
            {
                Log.Error(Constants.LOGGING_SOURCE, CoreResources.GraphExtensions_ErrorOccured, ex.Error.Message);
                throw;
            }
        }

        /// <summary>
        /// Adds members to a Microsoft 365 group
        /// </summary>
        /// <param name="groupId">Id of the Microsoft 365 group to add the members to</param>
        /// <param name="members">String array with the UPNs of the users that need to be added as members to the group</param>
        /// <param name="accessToken">The OAuth 2.0 Access Token to use for invoking the Microsoft Graph</param>
        /// <param name="removeExistingMembers">If true, all existing members will be removed and only those provided will become members. If false, existing members will remain and the ones provided will be added to the list with existing members.</param>
        /// <param name="retryCount">Number of times to retry the request in case of throttling</param>
        /// <param name="delay">Milliseconds to wait before retrying the request. The delay will be increased (doubled) every retry</param>
        /// <param name="azureEnvironment">Defines the Azure Cloud Deployment. This is used to determine the MS Graph EndPoint to call which differs per Azure Cloud deployments. Defaults to Production (graph.microsoft.com).</param>
        public static void AddUnifiedGroupMembers(string groupId, string[] members, string accessToken, bool removeExistingMembers = false, int retryCount = 10, int delay = 500, AzureEnvironment azureEnvironment = AzureEnvironment.Production)
        {
            if (String.IsNullOrEmpty(accessToken))
            {
                throw new ArgumentNullException(nameof(accessToken));
            }

            try
            {
                Task.Run(async () =>
                {
                    var graphClient = CreateGraphClient(accessToken, retryCount, delay, azureEnvironment);

                    await UpdateMembers(members, graphClient, groupId, removeExistingMembers);

                }).GetAwaiter().GetResult();
            }
            catch (ServiceException ex)
            {
                Log.Error(Constants.LOGGING_SOURCE, CoreResources.GraphExtensions_ErrorOccured, ex.Error.Message);
                throw;
            }
        }

        /// <summary>
        /// Removes members from a Microsoft 365 group
        /// </summary>
        /// <param name="groupId">Id of the Microsoft 365 group to remove the members from</param>
        /// <param name="members">String array with the UPNs of the users that need to be removed as members from the group</param>
        /// <param name="accessToken">The OAuth 2.0 Access Token to use for invoking the Microsoft Graph</param>
        /// <param name="retryCount">Number of times to retry the request in case of throttling</param>
        /// <param name="delay">Milliseconds to wait before retrying the request. The delay will be increased (doubled) every retry</param>
        /// <param name="azureEnvironment">Defines the Azure Cloud Deployment. This is used to determine the MS Graph EndPoint to call which differs per Azure Cloud deployments. Defaults to Production (graph.microsoft.com).</param>
        public static void RemoveUnifiedGroupMembers(string groupId, string[] members, string accessToken, int retryCount = 10, int delay = 500, AzureEnvironment azureEnvironment = AzureEnvironment.Production)
        {
            if (String.IsNullOrEmpty(accessToken))
            {
                throw new ArgumentNullException(nameof(accessToken));
            }

            try
            {
                Task.Run(async () =>
                {
                    var graphClient = CreateGraphClient(accessToken, retryCount, delay, azureEnvironment);

                    foreach (var m in members)
                    {
                        // Search for the user object
                        var memberQuery = await graphClient.Users
                            .Request()
                            .Filter($"userPrincipalName eq '{Uri.EscapeDataString(m.Replace("'", "''"))}'")
                            .GetAsync();

                        var member = memberQuery.FirstOrDefault();

                        if (member != null)
                        {
                            try
                            {
                                // If it is not in the list of current members, just remove it
                                await graphClient.Groups[groupId].Members[member.Id].Reference.Request().DeleteAsync();
                            }
                            catch (ServiceException ex)
                            {
                                if (ex.Error.Code == "Request_BadRequest")
                                {
                                    // Skip any failing removal
                                }
                                else
                                {
                                    throw ex;
                                }
                            }
                        }
                    }

                }).GetAwaiter().GetResult();
            }
            catch (ServiceException ex)
            {
                Log.Error(Constants.LOGGING_SOURCE, CoreResources.GraphExtensions_ErrorOccured, ex.Error.Message);
                throw;
            }
        }

        /// <summary>
        /// Removes owners from a Microsoft 365 group
        /// </summary>
        /// <param name="groupId">Id of the Microsoft 365 group to remove the owners from</param>
        /// <param name="owners">String array with the UPNs of the users that need to be removed as owners from the group</param>
        /// <param name="accessToken">The OAuth 2.0 Access Token to use for invoking the Microsoft Graph</param>
        /// <param name="retryCount">Number of times to retry the request in case of throttling</param>
        /// <param name="delay">Milliseconds to wait before retrying the request. The delay will be increased (doubled) every retry</param>
        /// <param name="azureEnvironment">Defines the Azure Cloud Deployment. This is used to determine the MS Graph EndPoint to call which differs per Azure Cloud deployments. Defaults to Production (graph.microsoft.com).</param>
        public static void RemoveUnifiedGroupOwners(string groupId, string[] owners, string accessToken, int retryCount = 10, int delay = 500, AzureEnvironment azureEnvironment = AzureEnvironment.Production)
        {
            if (String.IsNullOrEmpty(accessToken))
            {
                throw new ArgumentNullException(nameof(accessToken));
            }

            try
            {
                Task.Run(async () =>
                {
                    var graphClient = CreateGraphClient(accessToken, retryCount, delay, azureEnvironment);

                    foreach (var m in owners)
                    {
                        // Search for the user object
                        var memberQuery = await graphClient.Users
                            .Request()
                            .Filter($"userPrincipalName eq '{Uri.EscapeDataString(m.Replace("'", "''"))}'")
                            .GetAsync();

                        var member = memberQuery.FirstOrDefault();

                        if (member != null)
                        {
                            try
                            {
                                // If it is not in the list of current owners, just remove it
                                await graphClient.Groups[groupId].Owners[member.Id].Reference.Request().DeleteAsync();
                            }
                            catch (ServiceException ex)
                            {
                                if (ex.Error.Code == "Request_BadRequest")
                                {
                                    // Skip any failing removal
                                }
                                else
                                {
                                    throw ex;
                                }
                            }
                        }
                    }

                }).GetAwaiter().GetResult();
            }
            catch (ServiceException ex)
            {
                Log.Error(Constants.LOGGING_SOURCE, CoreResources.GraphExtensions_ErrorOccured, ex.Error.Message);
                throw;
            }
        }

        /// <summary>
        /// Removes all owners of a Microsoft 365 group
        /// </summary>
        /// <param name="groupId">Id of the Microsoft 365 group to remove all the current owners of</param>
        /// <param name="accessToken">The OAuth 2.0 Access Token to use for invoking the Microsoft Graph</param>
        /// <param name="retryCount">Number of times to retry the request in case of throttling</param>
        /// <param name="delay">Milliseconds to wait before retrying the request. The delay will be increased (doubled) every retry</param>
        /// <param name="azureEnvironment">Defines the Azure Cloud Deployment. This is used to determine the MS Graph EndPoint to call which differs per Azure Cloud deployments. Defaults to Production (graph.microsoft.com).</param>
        public static void ClearUnifiedGroupOwners(string groupId, string accessToken, int retryCount = 10, int delay = 500, AzureEnvironment azureEnvironment = AzureEnvironment.Production)
        {
            if (String.IsNullOrEmpty(accessToken))
            {
                throw new ArgumentNullException(nameof(accessToken));
            }

            try
            {
                var currentOwners = GetUnifiedGroupOwners(new UnifiedGroupEntity { GroupId = groupId }, accessToken, retryCount, delay, azureEnvironment);
                RemoveUnifiedGroupOwners(groupId, currentOwners.Select(o => o.UserPrincipalName).ToArray(), accessToken, retryCount, delay, azureEnvironment);
            }
            catch (ServiceException ex)
            {
                Log.Error(Constants.LOGGING_SOURCE, CoreResources.GraphExtensions_ErrorOccured, ex.Error.Message);
                throw;
            }
        }

        /// <summary>
        /// Removes all members of a Microsoft 365 group
        /// </summary>
        /// <param name="groupId">Id of the Microsoft 365 group to remove all the current members of</param>
        /// <param name="accessToken">The OAuth 2.0 Access Token to use for invoking the Microsoft Graph</param>
        /// <param name="retryCount">Number of times to retry the request in case of throttling</param>
        /// <param name="delay">Milliseconds to wait before retrying the request. The delay will be increased (doubled) every retry</param>
        /// <param name="azureEnvironment">Defines the Azure Cloud Deployment. This is used to determine the MS Graph EndPoint to call which differs per Azure Cloud deployments. Defaults to Production (graph.microsoft.com).</param>
        public static void ClearUnifiedGroupMembers(string groupId, string accessToken, int retryCount = 10, int delay = 500, AzureEnvironment azureEnvironment = AzureEnvironment.Production)
        {
            if (String.IsNullOrEmpty(accessToken))
            {
                throw new ArgumentNullException(nameof(accessToken));
            }

            try
            {
                var currentMembers = GetUnifiedGroupMembers(new UnifiedGroupEntity { GroupId = groupId }, accessToken, retryCount, delay, azureEnvironment);
                RemoveUnifiedGroupMembers(groupId, currentMembers.Select(o => o.UserPrincipalName).ToArray(), accessToken, retryCount, delay, azureEnvironment);
            }
            catch (ServiceException ex)
            {
                Log.Error(Constants.LOGGING_SOURCE, CoreResources.GraphExtensions_ErrorOccured, ex.Error.Message);
                throw;
            }
        }

        /// <summary>
        /// Returns all the Owners of an Office 365 group.
        /// </summary>
        /// <param name="group">The Office 365 group object of type UnifiedGroupEntity</param>
        /// <param name="accessToken">The OAuth 2.0 Access Token to use for invoking the Microsoft Graph</param>
        /// <param name="retryCount">Number of times to retry the request in case of throttling</param>
        /// <param name="delay">Milliseconds to wait before retrying the request. The delay will be increased (doubled) every retry</param>
        /// <param name="azureEnvironment">Defines the Azure Cloud Deployment. This is used to determine the MS Graph EndPoint to call which differs per Azure Cloud deployments. Defaults to Production (graph.microsoft.com).</param>
        /// <returns>Owners of an Office 365 group as a list of UnifiedGroupUser entity</returns>
        public static List<UnifiedGroupUser> GetUnifiedGroupOwners(UnifiedGroupEntity group, string accessToken, int retryCount = 10, int delay = 500, AzureEnvironment azureEnvironment = AzureEnvironment.Production)
        {
            List<UnifiedGroupUser> unifiedGroupUsers = null;
            List<User> unifiedGroupGraphUsers = null;
            IGroupOwnersCollectionWithReferencesPage groupUsers = null;

            if (String.IsNullOrEmpty(accessToken))
            {
                throw new ArgumentNullException(nameof(accessToken));
            }

            try
            {
                var result = Task.Run(async () =>
                {
                    var graphClient = CreateGraphClient(accessToken, retryCount, delay, azureEnvironment);

                    // Get the owners of an Office 365 group.
                    groupUsers = await graphClient.Groups[group.GroupId].Owners.Request().GetAsync();
                    if (groupUsers.CurrentPage != null && groupUsers.CurrentPage.Count > 0)
                    {
                        unifiedGroupGraphUsers = new List<User>();
                        GenerateGraphUserCollection(groupUsers.CurrentPage, unifiedGroupGraphUsers);
                    }

                    // Retrieve users when the results are paged.
                    while (groupUsers.NextPageRequest != null)
                    {
                        groupUsers = groupUsers.NextPageRequest.GetAsync().GetAwaiter().GetResult();
                        if (groupUsers.CurrentPage != null && groupUsers.CurrentPage.Count > 0)
                        {
                            GenerateGraphUserCollection(groupUsers.CurrentPage, unifiedGroupGraphUsers);
                        }
                    }

                    // Create the collection of type OfficeDevPnP 'UnifiedGroupUser' after all users are retrieved, including paged data.
                    if (unifiedGroupGraphUsers != null && unifiedGroupGraphUsers.Count > 0)
                    {
                        unifiedGroupUsers = new List<UnifiedGroupUser>();
                        foreach (User usr in unifiedGroupGraphUsers)
                        {
                            UnifiedGroupUser groupUser = new UnifiedGroupUser
                            {
                                Id = usr.Id,
                                UserPrincipalName = usr.UserPrincipalName != null ? usr.UserPrincipalName : string.Empty,
                                DisplayName = usr.DisplayName != null ? usr.DisplayName : string.Empty,
                                GivenName = usr.GivenName != null ? usr.GivenName : string.Empty,
                                Surname = usr.Surname != null ? usr.Surname : string.Empty,
                                Email = usr.Mail != null ? usr.Mail : string.Empty,
                                MobilePhone = usr.MobilePhone != null ? usr.DisplayName : string.Empty,
                                PreferredLanguage = usr.PreferredLanguage != null ? usr.PreferredLanguage : string.Empty,
                                JobTitle = usr.JobTitle != null ? usr.DisplayName : string.Empty,
                                BusinessPhones = usr.BusinessPhones != null ? usr.BusinessPhones.ToArray() : null
                            };
                            unifiedGroupUsers.Add(groupUser);
                        }
                    }
                    return unifiedGroupUsers;

                }).GetAwaiter().GetResult();
            }
            catch (ServiceException ex)
            {
                Log.Error(Constants.LOGGING_SOURCE, CoreResources.GraphExtensions_ErrorOccured, ex.Error.Message);
                throw;
            }
            return unifiedGroupUsers;
        }

        /// <summary>
        /// Helper method. Generates a collection of Microsoft.Graph.User entity from directory objects.
        /// </summary>
        /// <param name="page"></param>
        /// <param name="unifiedGroupGraphUsers"></param>
        /// <returns>Returns a collection of Microsoft.Graph.User entity</returns>
        private static List<User> GenerateGraphUserCollection(IList<DirectoryObject> page, List<User> unifiedGroupGraphUsers)
        {
            // Create a collection of Microsoft.Graph.User type
            foreach (User usr in page)
            {
                if (usr != null)
                {
                    unifiedGroupGraphUsers.Add(usr);
                }
            }

            return unifiedGroupGraphUsers;
        }

        /// <summary>
        /// Helper method. Generates a neseted collection of Microsoft.Graph.User entity from directory objects.
        /// </summary>
        /// <param name="page"></param>
        /// <param name="unifiedGroupGraphUsers"></param>
        /// <param name="unifiedGroupUsers"></param>
        /// <param name="accessToken"></param>
        /// <param name="azureEnvironment">Defines the Azure Cloud Deployment. This is used to determine the MS Graph EndPoint to call which differs per Azure Cloud deployments. Defaults to Production (graph.microsoft.com).</param>
        /// <returns></returns>
        private static List<User> GenerateNestedGraphUserCollection(IList<DirectoryObject> page, List<User> unifiedGroupGraphUsers, List<UnifiedGroupUser> unifiedGroupUsers, string accessToken, AzureEnvironment azureEnvironment = AzureEnvironment.Production)
        {
            // Create a collection of Microsoft.Graph.User type
            foreach (var usr in page)
            {

                if (usr != null)
                {
                    if (usr.GetType() == typeof(User))
                    {
                        unifiedGroupGraphUsers.Add((User)usr);
                    }
                }
            }

            //Get groups within the group and users in that group
            List<Group> unifiedGroupGraphGroups = new List<Group>();
            GenerateGraphGroupCollection(page, unifiedGroupGraphGroups);
            foreach (Group unifiedGroupGraphGroup in unifiedGroupGraphGroups)
            {
                var grp = GetUnifiedGroup(unifiedGroupGraphGroup.Id, accessToken, azureEnvironment: azureEnvironment);
                unifiedGroupUsers.AddRange(GetUnifiedGroupMembers(grp, accessToken));
            }

            return unifiedGroupGraphUsers;
        }

        /// <summary>
        /// Helper method. Generates a collection of Microsoft.Graph.Group entity from directory objects.
        /// </summary>
        /// <param name="page"></param>
        /// <param name="unifiedGroupGraphGroups"></param>
        /// <returns></returns>
        private static List<Group> GenerateGraphGroupCollection(IList<DirectoryObject> page, List<Group> unifiedGroupGraphGroups)
        {
            // Create a collection of Microsoft.Graph.Group type
            foreach (var grp in page)
            {

                if (grp != null)
                {
                    if (grp.GetType() == typeof(Group))
                    {
                        unifiedGroupGraphGroups.Add((Group)grp);
                    }
                }
            }

            return unifiedGroupGraphGroups;
        }

        /// <summary>
        /// Helper method. Generates a collection of Microsoft.Graph.User entity from string array
        /// </summary>
        /// <param name="graphClient">Graph service client</param>
        /// <param name="groupUsers">String array of users</param>
        /// <returns></returns>

        private static List<User> GetUsers(GraphServiceClient graphClient, string[] groupUsers)
        {
            if (groupUsers == null || groupUsers.Length == 0)
            {
                return new List<User>();
            }

            var result = Task.Run(async () =>
            {
                var usersResult = new List<User>();
                foreach (string groupUser in groupUsers)
                {
                    try
                    {
                        // Search for the user object
                        IGraphServiceUsersCollectionPage userQuery = await graphClient.Users
                                            .Request()
                                            .Select("Id")
                                            .Filter($"userPrincipalName eq '{Uri.EscapeDataString(groupUser.Replace("'", "''"))}'")
                                            .GetAsync();

                        User user = userQuery.FirstOrDefault();
                        if (user != null)
                        {
                            usersResult.Add(user);
                        }
                    }
                    catch (ServiceException)
                    {
                        // skip, group provisioning shouldnt stop because of error in user object
                    }
                }
                return usersResult;
            }).GetAwaiter().GetResult();
            return result;
        }

        /// <summary>
        /// Returns the classification value of an Office 365 Group.
        /// </summary>
        /// <param name="groupId">ID of the unified Group</param>
        /// <param name="accessToken">The OAuth 2.0 Access Token to use for invoking the Microsoft Graph</param>
        /// <param name="azureEnvironment">Defines the Azure Cloud Deployment. This is used to determine the MS Graph EndPoint to call which differs per Azure Cloud deployments. Defaults to Production (graph.microsoft.com).</param>
        /// <returns>Classification value of a Unified group</returns>
        public static string GetGroupClassification(string groupId, string accessToken, AzureEnvironment azureEnvironment = AzureEnvironment.Production)
        {
            if (String.IsNullOrEmpty(groupId))
            {
                throw new ArgumentNullException(nameof(groupId));
            }

            if (String.IsNullOrEmpty(accessToken))
            {
                throw new ArgumentNullException(nameof(accessToken));
            }

            string classification = string.Empty;

            try
            {
                string getGroupUrl = $"{GraphHttpClient.GetGraphEndPointUrl(azureEnvironment)}groups/{groupId}";

                var getGroupResult = GraphHttpClient.MakeGetRequestForString(
                    getGroupUrl,
                    accessToken: accessToken);

                JObject groupObject = JObject.Parse(getGroupResult);

                if (groupObject["classification"] != null)
                {
                    classification = Convert.ToString(groupObject["classification"]);
                }

            }
            catch (ServiceException e)
            {
                classification = e.Error.Message;
            }

            return classification;
        }

        /// <summary>
        /// Does this group have a Teams team?
        /// </summary>
        /// <param name="groupId">Id of the group to check</param>
        /// <param name="accessToken">Access token with scope Group.Read.All</param>
        /// <param name="azureEnvironment">Defines the Azure Cloud Deployment. This is used to determine the MS Graph EndPoint to call which differs per Azure Cloud deployments. Defaults to Production (graph.microsoft.com).</param>
        /// <returns>True if there's a Teams linked to this group</returns>
        public static bool HasTeamsTeam(string groupId, string accessToken, AzureEnvironment azureEnvironment = AzureEnvironment.Production)
        {
            if (String.IsNullOrEmpty(groupId))
            {
                throw new ArgumentNullException(nameof(groupId));
            }

            if (String.IsNullOrEmpty(accessToken))
            {
                throw new ArgumentNullException(nameof(accessToken));
            }

            bool hasTeamsTeam = false;

            try
            {
                groupId = groupId.ToLower();

                string getGroupsInfo = $"{GraphHttpClient.GetGraphEndPointUrl(azureEnvironment)}groups/{groupId}?$select=resourceProvisioningOptions";

                var getGroupResult = GraphHttpClient.MakeGetRequestForString(
                    getGroupsInfo,
                    accessToken: accessToken);

                using (var jsonDocument = JsonDocument.Parse(getGroupResult))
                {
                    var rootElement = jsonDocument.RootElement;
                    if (rootElement.TryGetProperty("resourceProvisioningOptions", out JsonElement resourceProvisioningOptionsElement))
                    {
                        return resourceProvisioningOptionsElement.EnumerateArray().FirstOrDefault(p => p.ValueEquals("Team")).ValueKind != JsonValueKind.Undefined;
                    }
                }
            }
            catch (ServiceException ex)
            {
                Log.Error(Constants.LOGGING_SOURCE, CoreResources.GraphExtensions_ErrorOccured, ex.Error.Message);
                throw;
            }

            return hasTeamsTeam;
        }

        /// <summary>
        /// Creates a team associated with an Office 365 group
        /// </summary>
        /// <param name="groupId">The ID of the Office 365 Group</param>
        /// <param name="accessToken">The OAuth 2.0 Access Token to use for invoking the Microsoft Graph</param>
        /// <param name="timeoutSeconds">Time to wait till Team is created. Default is 300 seconds (5 mins)</param>
        /// <param name="azureEnvironment">Defines the Azure Cloud Deployment. This is used to determine the MS Graph EndPoint to call which differs per Azure Cloud deployments. Defaults to Production (graph.microsoft.com).</param>
        /// <returns></returns>
        public static async Task CreateTeam(string groupId, string accessToken, int timeoutSeconds = 300, AzureEnvironment azureEnvironment = AzureEnvironment.Production)
        {
            if (string.IsNullOrEmpty(groupId))
            {
                throw new ArgumentNullException(nameof(groupId));
            }
            if (string.IsNullOrEmpty(accessToken))
            {
                throw new ArgumentNullException(nameof(accessToken));
            }

            var createTeamEndPoint = $"{GraphHttpClient.GetGraphEndPointUrl(azureEnvironment)}groups/{groupId}/team";
            bool wait = true;
            int iterations = 0;
            while (wait)
            {
                iterations++;
                try
                {
                    await Task.Run(() =>
                    {
                        var teamid = HttpHelper.MakePutRequestForString(createTeamEndPoint, new { }, "application/json", accessToken);
                        if (!string.IsNullOrEmpty(teamid))
                        {
                            wait = false;
                        }
                    });
                }
                catch (Exception ex)
                {
                    // Don't wait more than the requested timeout in seconds
                    if (iterations * 30 >= timeoutSeconds)
                    {
                        wait = false;
                        throw;
                    }
                    else
                    {
                        // In case of exception wait for 30 secs
                        Log.Error(Constants.LOGGING_SOURCE, CoreResources.GraphExtensions_ErrorOccured, ex.Message);
                        await Task.Delay(TimeSpan.FromSeconds(30));
                    }
                }
            }
        }

        /// <summary>
        /// Gets one deleted unified group based on its ID.
        /// </summary>
        /// <param name="groupId">The ID of the deleted group.</param>
        /// <param name="accessToken">Access token for accessing Microsoft Graph</param>
        /// <param name="azureEnvironment">Defines the Azure Cloud Deployment. This is used to determine the MS Graph EndPoint to call which differs per Azure Cloud deployments. Defaults to Production (graph.microsoft.com).</param>
        /// <returns>The unified group object of the deleted group that matches the provided ID.</returns>
        public static UnifiedGroupEntity GetDeletedUnifiedGroup(string groupId, string accessToken, AzureEnvironment azureEnvironment = AzureEnvironment.Production)
        {
            try
            {
                var response = HttpHelper.MakeGetRequestForString($"{GraphHttpClient.GetGraphEndPointUrl(azureEnvironment)}directory/deleteditems/microsoft.graph.group/{groupId}", accessToken);

                var group = JToken.Parse(response);

                var deletedGroup = new UnifiedGroupEntity
                {
                    GroupId = group["id"].ToString(),
                    Classification = group["classification"].ToString(),
                    Description = group["description"].ToString(),
                    DisplayName = group["displayName"].ToString(),
                    Mail = group["mail"].ToString(),
                    MailNickname = group["mailNickname"].ToString(),
                    Visibility = group["visibility"].ToString()
                };

                return deletedGroup;
            }
            catch (Exception e)
            {
                Log.Error(Constants.LOGGING_SOURCE, CoreResources.GraphExtensions_ErrorOccured, e.Message);
                throw;
            }
        }

        /// <summary>
        ///  Lists deleted unified groups.
        /// </summary>
        /// <param name="accessToken">Access token for accessing Microsoft Graph</param>
        /// <param name="azureEnvironment">Defines the Azure Cloud Deployment. This is used to determine the MS Graph EndPoint to call which differs per Azure Cloud deployments. Defaults to Production (graph.microsoft.com).</param>
        /// <returns>A list of unified group objects for the deleted groups.</returns>
        public static List<UnifiedGroupEntity> ListDeletedUnifiedGroups(string accessToken, AzureEnvironment azureEnvironment = AzureEnvironment.Production)
        {
            return ListDeletedUnifiedGroups(accessToken, null, null, azureEnvironment);
        }

        private static List<UnifiedGroupEntity> ListDeletedUnifiedGroups(string accessToken, List<UnifiedGroupEntity> deletedGroups, string nextPageUrl, AzureEnvironment azureEnvironment = AzureEnvironment.Production)
        {
            try
            {
                if (deletedGroups == null) deletedGroups = new List<UnifiedGroupEntity>();

                var requestUrl = nextPageUrl ?? $"{GraphHttpClient.GetGraphEndPointUrl(azureEnvironment, true)}directory/deleteditems/microsoft.graph.group?filter=groupTypes/Any(x:x eq 'Unified')";
                var response = JToken.Parse(HttpHelper.MakeGetRequestForString(requestUrl, accessToken));

                var groups = response["value"];

                foreach (var group in groups)
                {
                    var deletedGroup = new UnifiedGroupEntity
                    {
                        GroupId = group["id"].ToString(),
                        Classification = group["classification"].ToString(),
                        Description = group["description"].ToString(),
                        DisplayName = group["displayName"].ToString(),
                        Mail = group["mail"].ToString(),
                        MailNickname = group["mailNickname"].ToString(),
                        Visibility = group["visibility"].ToString()
                    };

                    deletedGroups.Add(deletedGroup);
                }

                // has paging?
                return response["@odata.nextLink"] != null ? ListDeletedUnifiedGroups(accessToken, deletedGroups, response["@odata.nextLink"].ToString(), azureEnvironment) : deletedGroups;
            }
            catch (Exception e)
            {
                Log.Error(Constants.LOGGING_SOURCE, CoreResources.GraphExtensions_ErrorOccured, e.Message);
                throw;
            }
        }

        /// <summary>
        /// Restores one deleted unified group based on its ID.
        /// </summary>
        /// <param name="groupId">The ID of the deleted group.</param>
        /// <param name="accessToken">Access token for accessing Microsoft Graph</param>
        /// <param name="azureEnvironment">Defines the Azure Cloud Deployment. This is used to determine the MS Graph EndPoint to call which differs per Azure Cloud deployments. Defaults to Production (graph.microsoft.com).</param>
        /// <returns></returns>
        public static void RestoreDeletedUnifiedGroup(string groupId, string accessToken, AzureEnvironment azureEnvironment = AzureEnvironment.Production)
        {
            try
            {
                HttpHelper.MakePostRequest($"{GraphHttpClient.GetGraphEndPointUrl(azureEnvironment)}directory/deleteditems/{groupId}/restore", contentType: "application/json", accessToken: accessToken);
            }
            catch (Exception e)
            {
                Log.Error(Constants.LOGGING_SOURCE, CoreResources.GraphExtensions_ErrorOccured, e.Message);
                throw;
            }
        }

        /// <summary>
        /// Permanently deletes one deleted unified group based on its ID.
        /// </summary>
        /// <param name="groupId">The ID of the deleted group.</param>
        /// <param name="accessToken">Access token for accessing Microsoft Graph</param>
        /// <param name="azureEnvironment">Defines the Azure Cloud Deployment. This is used to determine the MS Graph EndPoint to call which differs per Azure Cloud deployments. Defaults to Production (graph.microsoft.com).</param>
        /// <returns></returns>
        public static void PermanentlyDeleteUnifiedGroup(string groupId, string accessToken, AzureEnvironment azureEnvironment = AzureEnvironment.Production)
        {
            try
            {
                HttpHelper.MakeDeleteRequest($"{GraphHttpClient.GetGraphEndPointUrl(azureEnvironment)}v1.0/directory/deleteditems/{groupId}", accessToken);
            }
            catch (Exception e)
            {
                Log.Error(Constants.LOGGING_SOURCE, CoreResources.GraphExtensions_ErrorOccured, e.Message);
                throw;
            }
        }
    }
}
