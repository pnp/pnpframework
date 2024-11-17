using Microsoft.Graph;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using PnP.Framework.Diagnostics;
using PnP.Framework.Entities;
using PnP.Framework.Utilities;
using PnP.Framework.Utilities.Graph;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text.Json;
using System.Text.Json.Nodes;
using System.Threading.Tasks;

namespace PnP.Framework.Graph
{
    /// <summary>
    /// Class that deals with Azure Active Directory group CRUD operations.
    /// </summary>
    public static class GroupsUtility
    {
        private const int defaultRetryCount = 10;
        private const int defaultDelay = 500;

        /// <summary>
        ///  Creates a new GraphServiceClient instance using a custom PnPHttpProvider
        /// </summary>
        /// <param name="accessToken">The OAuth 2.0 Access Token to configure the HTTP bearer Authorization Header</param>
        /// <param name="retryCount">Number of times to retry the request in case of throttling</param>
        /// <param name="delay">Milliseconds to wait before retrying the request.</param>
        /// <param name="azureEnvironment">Azure environment to use, needed to get the correct Microsoft Graph URL</param>
        /// <returns></returns>
        private static GraphServiceClient CreateGraphClient(string accessToken, int retryCount = defaultRetryCount, int delay = defaultDelay, AzureEnvironment azureEnvironment = AzureEnvironment.Production)
        {
            // Creates a new GraphServiceClient instance using a custom PnPHttpProvider
            // which natively supports retry logic for throttled requests
            // Default are 10 retries with a base delay of 500ms
            var result = new GraphServiceClient($"{AuthenticationManager.GetGraphBaseEndPoint(azureEnvironment)}v1.0", new DelegateAuthenticationProvider(
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
        /// Creates a new Azure Active Directory Group
        /// </summary>
        /// <param name="displayName">The Display Name for the Azure Active Directory Group</param>
        /// <param name="description">The Description for the Azure Active Directory Group</param>
        /// <param name="mailNickname">The Mail Nickname for the Azure Active Directory Group</param>
        /// <param name="mailEnabled">Boolean indicating if the group will be mail enabled</param>
        /// <param name="securityEnabled">Boolean indicating if the group will be security enabled</param>
        /// <param name="accessToken">The OAuth 2.0 Access Token to use for invoking the Microsoft Graph</param>
        /// <param name="owners">A list of UPNs for group owners, if any</param>
        /// <param name="members">A list of UPNs for group members, if any</param>
        /// <param name="retryCount">Number of times to retry the request in case of throttling</param>
        /// <param name="delay">Milliseconds to wait before retrying the request. The delay will be increased (doubled) every retry</param>
        /// <param name="azureEnvironment">Azure environment to use, needed to get the correct Microsoft Graph URL</param>
        /// <returns>The just created Azure Active Directory Group</returns>
        public static GroupEntity CreateGroup(string displayName, string description, string mailNickname, bool mailEnabled, bool securityEnabled,
            string accessToken, string[] owners = null, string[] members = null, int retryCount = 10, int delay = 500, AzureEnvironment azureEnvironment = AzureEnvironment.Production)
        {
            GroupEntity result = null;

            if (String.IsNullOrEmpty(displayName))
            {
                throw new ArgumentNullException(nameof(displayName));
            }

            if (String.IsNullOrEmpty(mailNickname) && mailEnabled)
            {
                throw new ArgumentNullException(nameof(mailNickname));
            }

            if (String.IsNullOrEmpty(accessToken))
            {
                throw new ArgumentNullException(nameof(accessToken));
            }

            try
            {

                // Prepare the group resource object
                var newGroup = new Model.Group
                {
                    DisplayName = displayName,
                    Description = string.IsNullOrEmpty(description) ? null : description,
                    MailNickname = mailNickname,
                    MailEnabled = mailEnabled,
                    SecurityEnabled = securityEnabled,
                };

                if (owners != null && owners.Length > 0)
                {
                    var userIds = GetUserIds(accessToken, owners, retryCount, delay, azureEnvironment);
                    if (userIds != null && userIds.Count > 0)
                    {
                        newGroup.OwnersODataBind = userIds.Select(u => string.Format("{1}/users/{0}", u, GraphHttpClient.GetGraphEndPointUrl(azureEnvironment))).ToArray();
                    }
                }

                if (members != null && members.Length > 0)
                {
                    var userIds = GetUserIds(accessToken, members, retryCount, delay, azureEnvironment);
                    if (userIds != null && userIds.Count > 0)
                    {
                        newGroup.MembersODataBind = userIds.Select(u => string.Format("{1}/users/{0}", u, GraphHttpClient.GetGraphEndPointUrl(azureEnvironment))).ToArray();
                    }
                }

                // Create the group
                var requestUrl = $"{GraphHttpClient.GetGraphEndPointUrl(azureEnvironment)}groups";
                var responseAsString = HttpHelper.MakePostRequestForString(requestUrl, newGroup, accessToken: accessToken, retryCount: retryCount, delay: delay);
                return System.Text.Json.JsonSerializer.Deserialize<GroupEntity>(responseAsString);
            }
            catch (HttpRequestException ex)
            {
                Log.Error(Constants.LOGGING_SOURCE, CoreResources.GraphExtensions_ErrorOccured, ex.Message);
                throw;
            }
        }

        /// <summary>
        /// Updates the members of an Azure Active Directory Group
        /// </summary>
        /// <param name="members">UPNs of users that need to be added as a member to the group</param>
        /// <param name="graphClient">GraphClient instance to use to communicate with the Microsoft Graph</param>
        /// <param name="groupId">Id of the group which needs the owners added</param>
        /// <param name="removeOtherMembers">If set to true, all existing members which are not specified through <paramref name="members"/> will be removed as a member from the group</param>
        private static async Task UpdateMembers(string[] members, GraphServiceClient graphClient, string groupId, bool removeOtherMembers, string accessToken, int retryCount, int delay, AzureEnvironment azureEnvironment)
        {
            var userRequestUrl = $"{GraphHttpClient.GetGraphEndPointUrl(azureEnvironment)}users";
            var groupRequestUrl = $"{GraphHttpClient.GetGraphEndPointUrl(azureEnvironment)}groups/{groupId}";
            foreach (var m in members)
            {
                // Search for the user object
                string upn = Uri.EscapeDataString(m.Replace("'", "''"));
                var requestUrl = $"{userRequestUrl}?$filter=userPrincipalName eq '{upn}'&$select=id";
                var responseAsString = HttpHelper.MakeGetRequestForString(requestUrl, accessToken, retryCount: retryCount, delay: delay);
                var jsonNode = JsonNode.Parse(responseAsString);
                var userListString = jsonNode["value"];
                var user = userListString.AsArray().FirstOrDefault()?.Deserialize<Model.User>();

                if (user != null)
                {
                    try
                    {
                        // And if any, add it to the collection of group's members
                        var memberUrl = $"{groupRequestUrl}/members/{user.Id}/ref";
                        HttpHelper.MakePostRequest(memberUrl, accessToken, retryCount: retryCount, delay: delay);
                    }
                    catch (Exception ex) when (ex.Message.Contains("Request_BadRequest") &&
                            ex.Message.Contains("added object references already exist"))
                    {
                        // Skip any already existing member
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
                            // If it is not in the list of current members, just remove it
                            var memberUrl = $"{groupRequestUrl}/members/{member.Id}/ref";
                            HttpHelper.MakeDeleteRequest(memberUrl, accessToken, retryCount: retryCount, delay: delay);
                        }
                        catch (HttpResponseException ex) when (ex.StatusCode == 400)
                        {
                            // Skip any failing removal
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
        /// Updates the owners of an Azure Active Directory Group
        /// </summary>
        /// <param name="owners">UPNs of users that need to be added as a owner to the group</param>
        /// <param name="graphClient">GraphClient instance to use to communicate with the Microsoft Graph</param>
        /// <param name="groupId">Id of the group which needs the owners added</param>
        /// <param name="removeOtherOwners">If set to true, all existing owners which are not specified through <paramref name="owners"/> will be removed as an owner from the group</param>
        private static async Task UpdateOwners(string[] owners, GraphServiceClient graphClient, string groupId, bool removeOtherOwners, string accessToken, int retryCount, int delay, AzureEnvironment azureEnvironment)
        {
            var userRequestUrl = $"{GraphHttpClient.GetGraphEndPointUrl(azureEnvironment)}users";
            var groupRequestUrl = $"{GraphHttpClient.GetGraphEndPointUrl(azureEnvironment)}groups/{groupId}";
            foreach (var o in owners)
            {
                // Search for the user object
                string upn = Uri.EscapeDataString(o.Replace("'", "''"));
                var requestUrl = $"{userRequestUrl}?$filter=userPrincipalName eq '{upn}'&$select=id";
                var responseAsString = HttpHelper.MakeGetRequestForString(requestUrl, accessToken, retryCount: retryCount, delay: delay);
                var jsonNode = JsonNode.Parse(responseAsString);
                var userListString = jsonNode["value"];
                var user = userListString.AsArray().FirstOrDefault()?.Deserialize<Model.User>();

                if (user != null)
                {
                    try
                    {
                        // And if any, add it to the collection of group's owners
                        var memberUrl = $"{groupRequestUrl}/owners/{user.Id}/ref";
                        HttpHelper.MakePostRequest(memberUrl, accessToken, retryCount: retryCount, delay: delay);
                    }
                    catch (Exception ex) when (ex.Message.Contains("Request_BadRequest") &&
                            ex.Message.Contains("added object references already exist"))
                    {
                        // Skip any already existing member
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
                            var memberUrl = $"{groupRequestUrl}/owners/{owner.Id}/ref";
                            HttpHelper.MakeDeleteRequest(memberUrl, accessToken, retryCount: retryCount, delay: delay);
                        }
                        catch (HttpResponseException ex) when (ex.StatusCode == 400)
                        {
                            // Skip any failing removal
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
        /// Sets the visibility of an Azure Active Directory Group
        /// </summary>
        /// <param name="groupId">Id of the Azure Active Directory Group to set the visibility state for</param>
        /// <param name="accessToken">The OAuth 2.0 Access Token to use for invoking the Microsoft Graph</param>
        /// <param name="hideFromAddressLists">True if the group should not be displayed in certain parts of the Outlook UI: the Address Book, address lists for selecting message recipients, and the Browse Groups dialog for searching groups; otherwise, false. Default value is false.</param>
        /// <param name="hideFromOutlookClients">True if the group should not be displayed in Outlook clients, such as Outlook for Windows and Outlook on the web; otherwise, false. Default value is false.</param>
        /// <param name="azureEnvironment">Defines the Azure Cloud Deployment. This is used to determine the MS Graph EndPoint to call which differs per Azure Cloud deployments. Defaults to Production (graph.microsoft.com).</param>
        public static void SetGroupVisibility(string groupId, string accessToken, bool? hideFromAddressLists, bool? hideFromOutlookClients, AzureEnvironment azureEnvironment = AzureEnvironment.Production)
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
                string updateGroupUrl = $"{GraphHttpClient.GetGraphEndPointUrl(azureEnvironment)}groups/{groupId}";
                var groupRequest = new Model.GroupPatchModel
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
            catch (HttpResponseException ex)
            {
                Log.Error(Constants.LOGGING_SOURCE, CoreResources.GraphExtensions_ErrorOccured, ex.Message);
                throw;
            }
        }

        /// <summary>
        /// Updates the Azure Active Directory Group
        /// </summary>
        /// <param name="groupId">The ID of the Azure Active Directory Group</param>
        /// <param name="displayName">The Display Name for the Azure Active Directory Group</param>
        /// <param name="description">The Description for the Azure Active Directory Group</param>
        /// <param name="owners">A list of UPNs for group owners, if any, to be added to the group</param>
        /// <param name="members">A list of UPNs for group members, if any, to be added to the group</param>
        /// <param name="securityEnabled">Boolean indicating if the group is enabled for setting permissions</param>
        /// <param name="mailEnabled">Boolean indicating if the group is enabled for distributing mail</param>
        /// <param name="accessToken">The OAuth 2.0 Access Token to use for invoking the Microsoft Graph</param>
        /// <param name="retryCount">Number of times to retry the request in case of throttling</param>
        /// <param name="delay">Milliseconds to wait before retrying the request. The delay will be increased (doubled) every retry</param>
        /// <param name="azureEnvironment">Azure environment to use, needed to get the correct Microsoft Graph URL</param>
        /// <returns>Boolean indicating whether the Azure Active Directory Group has been updated or not</returns>
        public static bool UpdateGroup(string groupId,
            string accessToken, int retryCount = 10, int delay = 500,
            string displayName = null, string description = null, string[] owners = null, string[] members = null, bool? securityEnabled = null, bool? mailEnabled = null, AzureEnvironment azureEnvironment = AzureEnvironment.Production)
        {
            bool result;
            try
            {
                var groupRequestUrl = $"{GraphHttpClient.GetGraphEndPointUrl(azureEnvironment)}groups/{groupId}";
                // Use a synchronous model to invoke the asynchronous process
                result = Task.Run(async () =>
                {
                    var graphClient = CreateGraphClient(accessToken, retryCount, delay, azureEnvironment);

                    var responseAsString = HttpHelper.MakeGetRequestForString(groupRequestUrl, accessToken, retryCount: retryCount, delay: delay);
                    var groupJson = JsonNode.Parse(responseAsString);
                    var groupToUpdate = groupJson.Deserialize<Model.Group>();

                    // Workaround for the PATCH request, needed after update to Graph Library
                    var clonedGroup = new Model.Group
                    {
                        GroupId = groupToUpdate.GroupId,
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

                    // Check if we need to update owners
                    if (owners != null && owners.Length > 0)
                    {
                        // For each and every owner
                        await UpdateOwners(owners, graphClient, groupToUpdate.GroupId, true, accessToken, retryCount, delay, azureEnvironment);
                        updateGroup = true;
                    }

                    // Check if we need to update members
                    if (members != null && members.Length > 0)
                    {
                        // For each and every owner
                        await UpdateMembers(members, graphClient, groupToUpdate.GroupId, true, accessToken, retryCount, delay, azureEnvironment);
                        updateGroup = true;
                    }

                    // Check if we have to update the MailEnabled property
                    if (mailEnabled.HasValue && mailEnabled != groupToUpdate.MailEnabled)
                    {
                        clonedGroup.MailEnabled = mailEnabled.Value;
                        updateGroup = true;
                    }

                    // Check if we have to update the SecurityEnabled property
                    if (securityEnabled.HasValue && securityEnabled != groupToUpdate.SecurityEnabled)
                    {
                        clonedGroup.SecurityEnabled = securityEnabled.Value;
                        updateGroup = true;
                    }

                    // If the Group has to be updated, just do it
                    if (updateGroup)
                    {
                        var updatedGroup = HttpHelper.MakePatchRequestForString(groupRequestUrl, clonedGroup, accessToken, retryCount: retryCount, delay: delay);
                        groupUpdated = true;
                    }

                    #endregion

                    // If any of the previous update actions has been completed
                    return groupUpdated;

                }).GetAwaiter().GetResult();
            }
            catch (HttpResponseException ex)
            {
                Log.Error(Constants.LOGGING_SOURCE, CoreResources.GraphExtensions_ErrorOccured, ex.Message);
                throw;
            }
            return result;
        }

        /// <summary>
        /// Deletes an Azure Active Directory Group
        /// </summary>
        /// <param name="groupId">The ID of the Azure Active Directory Group</param>
        /// <param name="accessToken">The OAuth 2.0 Access Token to use for invoking the Microsoft Graph</param>
        /// <param name="retryCount">Number of times to retry the request in case of throttling</param>
        /// <param name="delay">Milliseconds to wait before retrying the request. The delay will be increased (doubled) every retry</param>
        /// <param name="azureEnvironment">Azure environment to use, needed to get the correct Microsoft Graph URL</param>
        public static void DeleteGroup(string groupId, string accessToken, int retryCount = 10, int delay = 500, AzureEnvironment azureEnvironment = AzureEnvironment.Production)
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
                var requestUrl = $"{GraphHttpClient.GetGraphEndPointUrl(azureEnvironment)}groups/{groupId}";
                HttpHelper.MakeDeleteRequest(requestUrl, accessToken, retryCount: retryCount, delay: delay);
            }
            catch (HttpResponseException ex)
            {
                Log.Error(Constants.LOGGING_SOURCE, CoreResources.GraphExtensions_ErrorOccured, ex.Message);
                throw;
            }
        }

        /// <summary>
        /// Get an Azure Active Directory Group by Id
        /// </summary>
        /// <param name="groupId">The ID of the Azure Active Directory Group</param>
        /// <param name="accessToken">The OAuth 2.0 Access Token to use for invoking the Microsoft Graph</param>
        /// <param name="retryCount">Number of times to retry the request in case of throttling</param>
        /// <param name="delay">Milliseconds to wait before retrying the request. The delay will be increased (doubled) every retry</param>
        /// <param name="azureEnvironment">Azure environment to use, needed to get the correct Microsoft Graph URL</param>
        /// <returns>Group instance if found</returns>
        public static GroupEntity GetGroup(string groupId, string accessToken, int retryCount = 10, int delay = 500, AzureEnvironment azureEnvironment = AzureEnvironment.Production)
        {
            var group = GetRawGroup(groupId, accessToken, retryCount, delay, azureEnvironment);
            return group.AsEntity();
        }

        internal static PnP.Framework.Graph.Model.Group GetRawGroup(string groupId, string accessToken, int retryCount = 10, int delay = 500, AzureEnvironment azureEnvironment = AzureEnvironment.Production)
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
                var requestUrl = $"{GraphHttpClient.GetGraphEndPointUrl(azureEnvironment)}groups/{groupId}";
                var responseAsString = HttpHelper.MakeGetRequestForString(requestUrl, accessToken, retryCount: retryCount, delay: delay);
                var groupJson = JsonNode.Parse(responseAsString);
                var group = groupJson.Deserialize<Model.Group>();
                return group;
            }
            catch (HttpResponseException ex)
            {
                Log.Error(Constants.LOGGING_SOURCE, CoreResources.GraphExtensions_ErrorOccured, ex.Message);
                throw;
            }
        }

        /// <summary>
        /// Returns all the Azure Active Directory Groups in the current Tenant based on a startIndex.
        /// </summary>
        /// <param name="accessToken">The OAuth 2.0 Access Token to use for invoking the Microsoft Graph</param>
        /// <param name="displayName">The DisplayName of the Azure Active Directory Group. Leave NULL if you do not want to filter by display name.</param>
        /// <param name="mailNickname">The MailNickname of the Azure Active Directory Group. Leave NULL if you do not want to filter by mail nickname.</param>
        /// <param name="startIndex">If not specified, method will start with the first group.</param>
        /// <param name="endIndex">If not specified, method will return all groups.</param>
        /// <param name="retryCount">Number of times to retry the request in case of throttling</param>
        /// <param name="delay">Milliseconds to wait before retrying the request. The delay will be increased (doubled) every retry</param>
        /// <param name="pageSize">Page size used for the individual requests to Micrsoft Graph. Defaults to 999 which is currently the maximum value.</param>
        /// <param name="azureEnvironment">Azure environment to use, needed to get the correct Microsoft Graph URL</param>
        /// <returns>List of GroupEntity objects</returns>
        public static List<GroupEntity> GetGroups(string accessToken,
            string displayName = null, string mailNickname = null,
            int startIndex = 0, int? endIndex = null,
            int retryCount = 10, int delay = 500, int pageSize = 999, AzureEnvironment azureEnvironment = AzureEnvironment.Production)
        {
            if (string.IsNullOrEmpty(accessToken))
            {
                throw new ArgumentNullException(nameof(accessToken));
            }

            List<GroupEntity> result = null;
            try
            {
                var requestUrl = $"{GraphHttpClient.GetGraphEndPointUrl(azureEnvironment)}groups?$top={pageSize}";
                var filter = string.Empty;
                filter += !string.IsNullOrEmpty(displayName) ? $"(DisplayName eq '{Uri.EscapeDataString(displayName.Replace("'", "''"))}')" : string.Empty;
                filter += !string.IsNullOrEmpty(mailNickname) ? $"(MailNickname eq '{Uri.EscapeDataString(mailNickname.Replace("'", "''"))}')" : string.Empty;

                if (!string.IsNullOrWhiteSpace(filter))
                {
                    requestUrl += $"&$filter={filter}";
                }

                List<GroupEntity> groups = new List<GroupEntity>();
                int currentIndex = 0;
                foreach (var g in GraphUtility.ReadPagedDataFromRequest<Model.Group>(requestUrl, accessToken, retryCount, delay))
                {
                    if (groups.Count > endIndex)
                    {
                        break;
                    }
                        if (currentIndex >= startIndex)
                        {
                            groups.Add(g.AsEntity());
                        }
                        currentIndex++;
                    }
                }
            catch (HttpResponseException ex)
            {
                Log.Error(Constants.LOGGING_SOURCE, CoreResources.GraphExtensions_ErrorOccured, ex.Message);
                throw;
            }
            return (result);
        }

        /// <summary>
        /// Returns all the Members of an Azure Active Directory group
        /// </summary>
        /// <param name="group">The Azure Active Directory group to return its members of</param>
        /// <param name="accessToken">The OAuth 2.0 Access Token to use for invoking the Microsoft Graph</param>
        /// <param name="retryCount">Number of times to retry the request in case of throttling</param>
        /// <param name="delay">Milliseconds to wait before retrying the request. The delay will be increased (doubled) every retry</param>
        /// <param name="azureEnvironment">Azure environment to use, needed to get the correct Microsoft Graph URL</param>
        /// <returns>Members of an Azure Active Directory group</returns>
        public static List<GroupUser> GetGroupMembers(GroupEntity group, string accessToken, int retryCount = 10, int delay = 500, AzureEnvironment azureEnvironment = AzureEnvironment.Production)
        {
            List<GroupUser> groupUsers = null;
            List<DirectoryObject> groupGraphUsers = null;
            IGroupMembersCollectionWithReferencesPage groupUsersCollection = null;

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

                    // Get the members of the group
                    groupUsersCollection = await graphClient.Groups[group.GroupId].Members.Request().GetAsync();
                    if (groupUsersCollection.CurrentPage != null && groupUsersCollection.CurrentPage.Count > 0)
                    {
                        groupGraphUsers = new List<DirectoryObject>();
                        groupGraphUsers.AddRange(groupUsersCollection.CurrentPage);
                        //GenerateGraphUserCollection(groupUsersCollection.CurrentPage, groupGraphUsers);
                    }

                    // Retrieve users when the results are paged.
                    while (groupUsersCollection.NextPageRequest != null)
                    {
                        groupUsersCollection = groupUsersCollection.NextPageRequest.GetAsync().GetAwaiter().GetResult();
                        if (groupUsersCollection.CurrentPage != null && groupUsersCollection.CurrentPage.Count > 0)
                        {
                            groupGraphUsers.AddRange(groupUsersCollection.CurrentPage);
                            //GenerateGraphUserCollection(groupUsersCollection.CurrentPage, groupGraphUsers);
                        }
                    }

                    // Create the collection of type OfficeDevPnP groupuser after all users are retrieved, including paged data.
                    if (groupGraphUsers != null && groupGraphUsers.Count > 0)
                    {
                        groupUsers = new List<GroupUser>();
                        foreach (DirectoryObject usr in groupGraphUsers)
                        {
                            switch(usr)
                            {
                                case Microsoft.Graph.User userType:
                                    groupUsers.Add(new GroupUser
                                    {
                                        UserPrincipalName = userType.UserPrincipalName != null ? userType.UserPrincipalName : string.Empty,
                                        DisplayName = userType.DisplayName != null ? userType.DisplayName : string.Empty,
                                        Type = Enums.GroupUserType.User
                                    });
                                break;

                                case Microsoft.Graph.Group groupType:
                                    groupUsers.Add(new GroupUser
                                    {
                                        UserPrincipalName = groupType.Id != null ? groupType.Id : string.Empty,
                                        DisplayName = groupType.DisplayName != null ? groupType.DisplayName : string.Empty,
                                        Type = Enums.GroupUserType.Group
                                    });
                                    break;
                            }

                        }
                    }
                    return groupUsers;

                }).GetAwaiter().GetResult();
            }
            catch (ServiceException ex)
            {
                Log.Error(Constants.LOGGING_SOURCE, CoreResources.GraphExtensions_ErrorOccured, ex.Error.Message);
                throw;
            }
            return groupUsers;
        }

        /// <summary>
        /// Adds owners to an Azure Active Directory group
        /// </summary>
        /// <param name="groupId">Id of the Azure Active Directory Group to add the owners to</param>
        /// <param name="owners">String array with the UPNs of the users that need to be added as owners to the group</param>
        /// <param name="accessToken">The OAuth 2.0 Access Token to use for invoking the Microsoft Graph</param>
        /// <param name="removeExistingOwners">If true, all existing owners will be removed and only those provided will become owners. If false, existing owners will remain and the ones provided will be added to the list with existing owners.</param>
        /// <param name="retryCount">Number of times to retry the request in case of throttling</param>
        /// <param name="delay">Milliseconds to wait before retrying the request. The delay will be increased (doubled) every retry</param>
        /// <param name="azureEnvironment">Azure environment to use, needed to get the correct Microsoft Graph URL</param>
        public static void AddGroupOwners(string groupId, string[] owners, string accessToken, bool removeExistingOwners = false, int retryCount = 10, int delay = 500, AzureEnvironment azureEnvironment = AzureEnvironment.Production)
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

                    await UpdateOwners(owners, graphClient, groupId, removeExistingOwners, accessToken, retryCount, delay, azureEnvironment);

                }).GetAwaiter().GetResult();
            }
            catch (ServiceException ex)
            {
                Log.Error(Constants.LOGGING_SOURCE, CoreResources.GraphExtensions_ErrorOccured, ex.Error.Message);
                throw;
            }
        }

        /// <summary>
        /// Adds members to an Azure Active Directory group
        /// </summary>
        /// <param name="groupId">Id of the Azure Active Directory group to add the members to</param>
        /// <param name="members">String array with the UPNs of the users that need to be added as members to the group</param>
        /// <param name="accessToken">The OAuth 2.0 Access Token to use for invoking the Microsoft Graph</param>
        /// <param name="removeExistingMembers">If true, all existing members will be removed and only those provided will become members. If false, existing members will remain and the ones provided will be added to the list with existing members.</param>
        /// <param name="retryCount">Number of times to retry the request in case of throttling</param>
        /// <param name="delay">Milliseconds to wait before retrying the request. The delay will be increased (doubled) every retry</param>
        /// <param name="azureEnvironment">Azure environment to use, needed to get the correct Microsoft Graph URL</param>
        public static void AddGroupMembers(string groupId, string[] members, string accessToken, bool removeExistingMembers = false, int retryCount = 10, int delay = 500, AzureEnvironment azureEnvironment = AzureEnvironment.Production)
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

                    await UpdateMembers(members, graphClient, groupId, removeExistingMembers, accessToken, retryCount, delay, azureEnvironment);

                }).GetAwaiter().GetResult();
            }
            catch (ServiceException ex)
            {
                Log.Error(Constants.LOGGING_SOURCE, CoreResources.GraphExtensions_ErrorOccured, ex.Error.Message);
                throw;
            }
        }

        /// <summary>
        /// Removes members from an Azure Active Directory group
        /// </summary>
        /// <param name="groupId">Id of the Azure Active Directory group to remove the members from</param>
        /// <param name="members">String array with the UPNs of the users that need to be removed as members from the group</param>
        /// <param name="accessToken">The OAuth 2.0 Access Token to use for invoking the Microsoft Graph</param>
        /// <param name="retryCount">Number of times to retry the request in case of throttling</param>
        /// <param name="delay">Milliseconds to wait before retrying the request. The delay will be increased (doubled) every retry</param>
        /// <param name="azureEnvironment">Azure environment to use, needed to get the correct Microsoft Graph URL</param>
        public static void RemoveGroupMembers(string groupId, string[] members, string accessToken, int retryCount = 10, int delay = 500, AzureEnvironment azureEnvironment = AzureEnvironment.Production)
        {
            if (String.IsNullOrEmpty(accessToken))
            {
                throw new ArgumentNullException(nameof(accessToken));
            }

            try
            {
                var userRequestUrl = $"{GraphHttpClient.GetGraphEndPointUrl(azureEnvironment)}users";
                var groupRequestUrl = $"{GraphHttpClient.GetGraphEndPointUrl(azureEnvironment)}groups/{groupId}";

                foreach (var m in members)
                {
                    // Search for the user object
                    string upn = Uri.EscapeDataString(m.Replace("'", "''"));
                    var requestUrl = $"{userRequestUrl}?$filter=userPrincipalName eq '{upn}'&$select=id";
                    var responseAsString = HttpHelper.MakeGetRequestForString(requestUrl, accessToken, retryCount: retryCount, delay: delay);
                    var jsonNode = JsonNode.Parse(responseAsString);
                    var userListString= jsonNode["value"];
                    var userId = userListString.AsArray().FirstOrDefault()?["id"];

                    if (userId != null)
                    {
                        try
                        {
                            // If it is not in the list of current members, just remove it
                            var deleteGroupMemberUrl = $"{groupRequestUrl}/members/{userId}/ref";
                            HttpHelper.MakeDeleteRequest(deleteGroupMemberUrl, accessToken, retryCount: retryCount, delay: delay);
                        }
                        catch (HttpResponseException ex) when (ex.StatusCode == 400)
                        {
                            // Skip any failing removal
                        }
                    }
                }
            }
            catch (HttpResponseException ex)
            {
                Log.Error(Constants.LOGGING_SOURCE, CoreResources.GraphExtensions_ErrorOccured, ex.Message);
                throw;
            }
        }

        /// <summary>
        /// Removes owners from an Azure Active Directory group
        /// </summary>
        /// <param name="groupId">Id of the Azure Active Directory group to remove the owners from</param>
        /// <param name="owners">String array with the UPNs of the users that need to be removed as owners from the group</param>
        /// <param name="accessToken">The OAuth 2.0 Access Token to use for invoking the Microsoft Graph</param>
        /// <param name="retryCount">Number of times to retry the request in case of throttling</param>
        /// <param name="delay">Milliseconds to wait before retrying the request. The delay will be increased (doubled) every retry</param>
        /// <param name="azureEnvironment">Azure environment to use, needed to get the correct Microsoft Graph URL</param>
        public static void RemoveGroupOwners(string groupId, string[] owners, string accessToken, int retryCount = 10, int delay = 500, AzureEnvironment azureEnvironment = AzureEnvironment.Production)
        {
            if (String.IsNullOrEmpty(accessToken))
            {
                throw new ArgumentNullException(nameof(accessToken));
            }

            try
            {
                var userRequestUrl = $"{GraphHttpClient.GetGraphEndPointUrl(azureEnvironment)}users";
                var groupRequestUrl = $"{GraphHttpClient.GetGraphEndPointUrl(azureEnvironment)}groups/{groupId}";

                foreach (var m in owners)
                {
                    // Search for the user object
                    string upn = Uri.EscapeDataString(m.Replace("'", "''"));
                    var requestUrl = $"{userRequestUrl}?$filter=userPrincipalName eq '{upn}'&$select=id";
                    var responseAsString = HttpHelper.MakeGetRequestForString(requestUrl, accessToken, retryCount: retryCount, delay: delay);
                    var jsonNode = JsonNode.Parse(responseAsString);
                    var userListString = jsonNode["value"];
                    var userId = userListString.AsArray().FirstOrDefault()?["id"];

                    if (userId != null)
                    {
                        try
                        {
                            // If it is not in the list of current owners, just remove it
                            var deleteGroupMemberUrl = $"{groupRequestUrl}/owners/{userId}/ref";
                            HttpHelper.MakeDeleteRequest(deleteGroupMemberUrl, accessToken, retryCount: retryCount, delay: delay);
                        }
                        catch (HttpResponseException ex) when (ex.StatusCode == 400)
                        {
                            // Skip any failing removal
                        }
                    }
                }
            }
            catch (HttpResponseException ex)
            {
                Log.Error(Constants.LOGGING_SOURCE, CoreResources.GraphExtensions_ErrorOccured, ex.Message);
                throw;
            }
        }

        /// <summary>
        /// Removes all owners of an Azure Active Directory group
        /// </summary>
        /// <param name="groupId">Id of the Azure Active Directory group to remove all the current owners of</param>
        /// <param name="accessToken">The OAuth 2.0 Access Token to use for invoking the Microsoft Graph</param>
        /// <param name="retryCount">Number of times to retry the request in case of throttling</param>
        /// <param name="delay">Milliseconds to wait before retrying the request. The delay will be increased (doubled) every retry</param>
        public static void ClearGroupOwners(string groupId, string accessToken, int retryCount = 10, int delay = 500)
        {
            if (String.IsNullOrEmpty(accessToken))
            {
                throw new ArgumentNullException(nameof(accessToken));
            }

            try
            {
                var currentOwners = GetGroupOwners(new GroupEntity { GroupId = groupId }, accessToken, retryCount, delay);
                RemoveGroupOwners(groupId, currentOwners.Select(o => o.UserPrincipalName).ToArray(), accessToken, retryCount, delay);
            }
            catch (ServiceException ex)
            {
                Log.Error(Constants.LOGGING_SOURCE, CoreResources.GraphExtensions_ErrorOccured, ex.Error.Message);
                throw;
            }
        }

        /// <summary>
        /// Removes all members of an Azure Active Directory group
        /// </summary>
        /// <param name="groupId">Id of the Azure Active Directory group to remove all the current members of</param>
        /// <param name="accessToken">The OAuth 2.0 Access Token to use for invoking the Microsoft Graph</param>
        /// <param name="retryCount">Number of times to retry the request in case of throttling</param>
        /// <param name="delay">Milliseconds to wait before retrying the request. The delay will be increased (doubled) every retry</param>
        public static void ClearGroupMembers(string groupId, string accessToken, int retryCount = 10, int delay = 500)
        {
            if (String.IsNullOrEmpty(accessToken))
            {
                throw new ArgumentNullException(nameof(accessToken));
            }

            try
            {
                var currentMembers = GetGroupMembers(new GroupEntity { GroupId = groupId }, accessToken, retryCount, delay);
                if (currentMembers == null) return;

                RemoveGroupMembers(groupId, currentMembers.Select(o => o.UserPrincipalName).ToArray(), accessToken, retryCount, delay);
            }
            catch (ServiceException ex)
            {
                Log.Error(Constants.LOGGING_SOURCE, CoreResources.GraphExtensions_ErrorOccured, ex.Error.Message);
                throw;
            }
        }

        /// <summary>
        /// Returns all the Owners of an Azure Active Directory group
        /// </summary>
        /// <param name="group">The Azure Active Directory group object</param>
        /// <param name="accessToken">The OAuth 2.0 Access Token to use for invoking the Microsoft Graph</param>
        /// <param name="retryCount">Number of times to retry the request in case of throttling</param>
        /// <param name="delay">Milliseconds to wait before retrying the request. The delay will be increased (doubled) every retry</param>
        /// <param name="azureEnvironment">Azure environment to use, needed to get the correct Microsoft Graph URL</param>
        /// <returns>Owners of an Azure Active Directory group</returns>
        public static List<GroupUser> GetGroupOwners(GroupEntity group, string accessToken, int retryCount = 10, int delay = 500, AzureEnvironment azureEnvironment = AzureEnvironment.Production)
        {
            List<GroupUser> groupUsers = null;
            List<User> groupGraphUsers = null;
            IGroupOwnersCollectionWithReferencesPage groupUsersCollection = null;

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
                    groupUsersCollection = await graphClient.Groups[group.GroupId].Owners.Request().GetAsync();
                    if (groupUsersCollection.CurrentPage != null && groupUsersCollection.CurrentPage.Count > 0)
                    {
                        groupGraphUsers = new List<User>();
                        GenerateGraphUserCollection(groupUsersCollection.CurrentPage, groupGraphUsers);
                    }

                    // Retrieve users when the results are paged.
                    while (groupUsersCollection.NextPageRequest != null)
                    {
                        groupUsersCollection = groupUsersCollection.NextPageRequest.GetAsync().GetAwaiter().GetResult();
                        if (groupUsersCollection.CurrentPage != null && groupUsersCollection.CurrentPage.Count > 0)
                        {
                            GenerateGraphUserCollection(groupUsersCollection.CurrentPage, groupGraphUsers);
                        }
                    }

                    // Create the collection of type OfficeDevPnP 'UnifiedGroupUser' after all users are retrieved, including paged data.
                    if (groupGraphUsers != null && groupGraphUsers.Count > 0)
                    {
                        groupUsers = new List<GroupUser>();
                        foreach (DirectoryObject usr in groupGraphUsers)
                        {
                            switch(usr)
                            {
                                case Microsoft.Graph.User userType:
                                    groupUsers.Add(new GroupUser
                                    {
                                        UserPrincipalName = userType.UserPrincipalName != null ? userType.UserPrincipalName : string.Empty,
                                        DisplayName = userType.DisplayName != null ? userType.DisplayName : string.Empty,
                                        Type = Enums.GroupUserType.User
                                    });
                                break;

                                case Microsoft.Graph.Group groupType:
                                    groupUsers.Add(new GroupUser
                                    {
                                        UserPrincipalName = groupType.Id != null ? groupType.Id : string.Empty,
                                        DisplayName = groupType.DisplayName != null ? groupType.DisplayName : string.Empty,
                                        Type = Enums.GroupUserType.Group
                                    });
                                    break;
                            }

                        }
                    }
                    return groupUsers;

                }).GetAwaiter().GetResult();
            }
            catch (ServiceException ex)
            {
                Log.Error(Constants.LOGGING_SOURCE, CoreResources.GraphExtensions_ErrorOccured, ex.Error.Message);
                throw;
            }
            return groupUsers;
        }

        /// <summary>
        /// Helper method. Generates a collection of Microsoft.Graph.User entity from directory objects.
        /// </summary>
        /// <param name="page"></param>
        /// <param name="groupGraphUsers"></param>
        /// <returns>Returns a collection of Microsoft.Graph.User entity</returns>
        private static List<User> GenerateGraphUserCollection(IList<DirectoryObject> page, List<User> groupGraphUsers)
        {
            // Create a collection of Microsoft.Graph.User type
            foreach (User usr in page)
            {
                if (usr != null)
                {
                    groupGraphUsers.Add(usr);
                }
            }

            return groupGraphUsers;
        }

        /// <summary>
        /// Helper method. Generates a collection of Microsoft.Graph.User entity from string array
        /// </summary>
        /// <param name="accessToken"></param>
        /// <param name="groupUsers">String array of users</param>
        /// <param name="retryCount"></param>
        /// <param name="delay"></param>
        /// <returns></returns>

        private static List<string> GetUserIds(string accessToken, string[] groupUsers, int retryCount, int delay, AzureEnvironment azureEnvironment)
        {
            if (groupUsers == null || groupUsers.Length == 0)
            {
                return new List<string>();
            }

            var usersResult = new List<string>();
            foreach (var groupUser in groupUsers)
            {
                try
                {
                    var requestUrl = $"{GraphHttpClient.GetGraphEndPointUrl(azureEnvironment)}users?$select=Id&$filter=userPrincipalName eq '{Uri.EscapeDataString(groupUser.Replace("'", "''"))}'";
                    var responseAsString = HttpHelper.MakeGetRequestForString(requestUrl, accessToken, retryCount: retryCount, delay: delay);

                    var jsonNode = JsonNode.Parse(responseAsString);
                    var usersArray = jsonNode["value"].AsArray();
                    var id = usersArray.FirstOrDefault()?["id"]?.GetValue<Guid>();

                    if (id != null)
                    {
                        usersResult.Add(id.Value.ToString());
                    }
                }
                catch (HttpResponseException)
                {
                    // skip, group provisioning shouldnt stop because of error in user object
                }
            }
            return usersResult;
        }

        /// <summary>
        /// Gets one deleted Azure Active Directory group based on its ID
        /// </summary>
        /// <param name="groupId">The ID of the deleted group.</param>
        /// <param name="accessToken">Access token for accessing Microsoft Graph</param>
        /// <param name="graphBaseUri">The Microsoft Graph URI to use</param>
        /// <returns>The unified group object of the deleted group that matches the provided ID.</returns>
        public static GroupEntity GetDeletedGroup(string groupId, string accessToken, Uri graphBaseUri = null)
        {
            try
            {
                if (graphBaseUri == null) 
                {
                    graphBaseUri = new Uri(GraphHelper.MicrosoftGraphBaseURI);
                }

                var response = HttpHelper.MakeGetRequestForString($"{graphBaseUri}v1.0/directory/deleteditems/microsoft.graph.group/{groupId}", accessToken);

                var group = JToken.Parse(response);

                var deletedGroup = new GroupEntity
                {
                    GroupId = group["id"].ToString(),
                    Description = group["description"].ToString(),
                    DisplayName = group["displayName"].ToString(),
                    Mail = group["mail"].ToString(),
                    MailNickname = group["mailNickname"].ToString(),
                    MailEnabled = group["mailEnabled"].ToString().Equals("true", StringComparison.InvariantCultureIgnoreCase),
                    SecurityEnabled = group["securityEnabled"].ToString().Equals("true", StringComparison.InvariantCultureIgnoreCase),
                    GroupTypes = group["GroupTypes"] != null ? new[] { group["GroupTypes"].ToString() } : null
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
        ///  Lists deleted Azure Active Directory groups
        /// </summary>
        /// <param name="accessToken">Access token for accessing Microsoft Graph</param>
        /// <param name="graphBaseUri">The Microsoft Graph URI to use</param>
        /// <returns>A list of Azure Active Directory group objects that have been deleted</returns>
        public static List<GroupEntity> ListDeletedGroups(string accessToken, Uri graphBaseUri = null)
        {
            return ListDeletedGroups(accessToken, null, null, graphBaseUri);
        }

        private static List<GroupEntity> ListDeletedGroups(string accessToken, List<GroupEntity> deletedGroups, string nextPageUrl, Uri graphBaseUri)
        {
            try
            {
                if (graphBaseUri == null) 
                {
                    graphBaseUri = new Uri(GraphHelper.MicrosoftGraphBaseURI);                
                }


                if (deletedGroups == null) deletedGroups = new List<GroupEntity>();

                var requestUrl = nextPageUrl ?? $"{graphBaseUri}beta/directory/deleteditems/microsoft.graph.group";
                var response = JToken.Parse(HttpHelper.MakeGetRequestForString(requestUrl, accessToken));

                var groups = response["value"];

                foreach (var group in groups)
                {
                    var deletedGroup = new GroupEntity
                    {
                        GroupId = group["id"].ToString(),
                        Description = group["description"].ToString(),
                        DisplayName = group["displayName"].ToString(),
                        Mail = group["mail"].ToString(),
                        MailNickname = group["mailNickname"].ToString(),
                        MailEnabled = group["mailEnabled"].ToString().Equals("true", StringComparison.InvariantCultureIgnoreCase),
                        SecurityEnabled = group["securityEnabled"].ToString().Equals("true", StringComparison.InvariantCultureIgnoreCase),
                        GroupTypes = group["GroupTypes"] != null ? new[] { group["GroupTypes"].ToString() } : null
                    };

                    deletedGroups.Add(deletedGroup);
                }

                // has paging?
                return response["@odata.nextLink"] != null ? ListDeletedGroups(accessToken, deletedGroups, response["@odata.nextLink"].ToString(), graphBaseUri) : deletedGroups;
            }
            catch (Exception e)
            {
                Log.Error(Constants.LOGGING_SOURCE, CoreResources.GraphExtensions_ErrorOccured, e.Message);
                throw;
            }
        }

        /// <summary>
        /// Restores one deleted Azure Active Directory group based on its ID
        /// </summary>
        /// <param name="groupId">The ID of the deleted group</param>
        /// <param name="accessToken">Access token for accessing Microsoft Graph</param>
        /// <param name="graphBaseUri">The Microsoft Graph URI to use</param>
        /// <returns></returns>
        public static void RestoreDeletedGroup(string groupId, string accessToken, Uri graphBaseUri = null)
        {
            try
            {
                if (graphBaseUri == null)
                {
                    graphBaseUri = new Uri(GraphHelper.MicrosoftGraphBaseURI);
                }

                HttpHelper.MakePostRequest($"{graphBaseUri}v1.0/directory/deleteditems/{groupId}/restore", contentType: "application/json", accessToken: accessToken);
            }
            catch (Exception e)
            {
                Log.Error(Constants.LOGGING_SOURCE, CoreResources.GraphExtensions_ErrorOccured, e.Message);
                throw;
            }
        }

        /// <summary>
        /// Permanently deletes one deleted Azure Active Directory group based on its ID
        /// </summary>
        /// <param name="groupId">The ID of the group to permanently delete</param>
        /// <param name="accessToken">Access token for accessing Microsoft Graph</param>
        /// <param name="graphBaseUri">The Microsoft Graph URI to use</param>
        /// <returns></returns>
        public static void PermanentlyDeleteGroup(string groupId, string accessToken, Uri graphBaseUri = null)
        {
            try
            {
                if (graphBaseUri == null)
                {
                    graphBaseUri = new Uri(GraphHelper.MicrosoftGraphBaseURI);
                }

                HttpHelper.MakeDeleteRequest($"{graphBaseUri}v1.0/directory/deleteditems/{groupId}", accessToken);
            }
            catch (Exception e)
            {
                Log.Error(Constants.LOGGING_SOURCE, CoreResources.GraphExtensions_ErrorOccured, e.Message);
                throw;
            }
        }
    }
}
