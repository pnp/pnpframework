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
using System.Net.Http.Headers;
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
                // Use a synchronous model to invoke the asynchronous process
                result = Task.Run(async () =>
                {
                    var group = new GroupEntity();

                    var graphClient = CreateGraphClient(accessToken, retryCount, delay, azureEnvironment);

                    // Prepare the group resource object
                    var newGroup = new GroupExtended
                    {
                        DisplayName = displayName,
                        Description = string.IsNullOrEmpty(description) ? null : description,
                        MailNickname = mailNickname,
                        MailEnabled = mailEnabled,
                        SecurityEnabled = securityEnabled
                    };

                    if (owners != null && owners.Length > 0)
                    {
                        var users = GetUsers(graphClient, owners);
                        if (users != null && users.Count > 0)
                        {
                            newGroup.OwnersODataBind = users.Select(u => string.Format("{1}/users/{0}", u.Id, graphClient.BaseUrl)).ToArray();
                        }
                    }

                    if (members != null && members.Length > 0)
                    {
                        var users = GetUsers(graphClient, members);
                        if (users != null && users.Count > 0)
                        {
                            newGroup.MembersODataBind = users.Select(u => string.Format("{1}/users/{0}", u.Id, graphClient.BaseUrl)).ToArray();
                        }
                    }

                    // Create the group
                    Microsoft.Graph.Group addedGroup = await graphClient.Groups.Request().AddAsync(newGroup);

                    if (addedGroup != null)
                    {
                        group.DisplayName = addedGroup.DisplayName;
                        group.Description = addedGroup.Description;
                        group.GroupId = addedGroup.Id;
                        group.Mail = addedGroup.Mail;
                        group.MailNickname = addedGroup.MailNickname;
                        group.MailEnabled = addedGroup.MailEnabled;
                        group.SecurityEnabled = addedGroup.SecurityEnabled;
                        group.GroupTypes = addedGroup.GroupTypes != null ? addedGroup.GroupTypes.ToArray() : null;
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
        /// Updates the members of an Azure Active Directory Group
        /// </summary>
        /// <param name="members">UPNs of users that need to be added as a member to the group</param>
        /// <param name="graphClient">GraphClient instance to use to communicate with the Microsoft Graph</param>
        /// <param name="groupId">Id of the group which needs the owners added</param>
        /// <param name="removeOtherMembers">If set to true, all existing members which are not specified through <paramref name="members"/> will be removed as a member from the group</param>
        private static async Task UpdateMembers(string[] members, GraphServiceClient graphClient, string groupId, bool removeOtherMembers)
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
        /// Updates the owners of an Azure Active Directory Group
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

                    // Check if we have to update the MailEnabled property
                    if (groupToUpdate.MailEnabled.HasValue)
                    {
                        clonedGroup.MailEnabled = groupToUpdate.MailEnabled.Value;
                        updateGroup = true;
                    }

                    // Check if we have to update the SecurityEnabled property
                    if (groupToUpdate.SecurityEnabled.HasValue)
                    {
                        clonedGroup.SecurityEnabled = groupToUpdate.SecurityEnabled.Value;
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

                    // If any of the previous update actions has been completed
                    return groupUpdated;

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
            if (string.IsNullOrEmpty(groupId))
            {
                throw new ArgumentNullException(nameof(groupId));
            }

            if (string.IsNullOrEmpty(accessToken))
            {
                throw new ArgumentNullException(nameof(accessToken));
            }

            GroupEntity result = null;
            try
            {
                // Use a synchronous model to invoke the asynchronous process
                result = Task.Run(async () =>
                {
                    GroupEntity group = null;

                    var graphClient = CreateGraphClient(accessToken, retryCount, delay, azureEnvironment);

                    var g = await graphClient.Groups[groupId].Request().GetAsync();

                    group = new GroupEntity
                    {
                        GroupId = g.Id,
                        DisplayName = g.DisplayName,
                        Description = g.Description,
                        Mail = g.Mail,
                        MailNickname = g.MailNickname,
                        MailEnabled = g.MailEnabled,
                        SecurityEnabled = g.SecurityEnabled,
                        GroupTypes = g.GroupTypes != null ? g.GroupTypes.ToArray() : null
                    };

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
                // Use a synchronous model to invoke the asynchronous process
                result = Task.Run(async () =>
                {
                    List<GroupEntity> groups = new List<GroupEntity>();

                    var graphClient = CreateGraphClient(accessToken, retryCount, delay, azureEnvironment);

                    // Apply the DisplayName filter, if any
                    var displayNameFilter = !string.IsNullOrEmpty(displayName) ? $"(DisplayName eq '{Uri.EscapeDataString(displayName.Replace("'", "''"))}')" : string.Empty;
                    var mailNicknameFilter = !string.IsNullOrEmpty(mailNickname) ? $"(MailNickname eq '{Uri.EscapeDataString(mailNickname.Replace("'", "''"))}')" : string.Empty;

                    var pagedGroups = await graphClient.Groups
                        .Request()
                        .Filter($"{displayNameFilter}{(!string.IsNullOrEmpty(displayNameFilter) && !string.IsNullOrEmpty(mailNicknameFilter) ? " and " : "")}{mailNicknameFilter}")
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
                                var group = new GroupEntity
                                {
                                    GroupId = g.Id,
                                    DisplayName = g.DisplayName,
                                    Description = g.Description,
                                    Mail = g.Mail,
                                    MailNickname = g.MailNickname,
                                    MailEnabled = g.MailEnabled,
                                    SecurityEnabled = g.SecurityEnabled,
                                    GroupTypes = g.GroupTypes != null ? g.GroupTypes.ToArray() : null
                                };

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
                            // If it is not in the list of current owners, just remove it
                            await graphClient.Groups[groupId].Owners[member.Id].Reference.Request().DeleteAsync();
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
        /// Helper method. Generates a neseted collection of Microsoft.Graph.User entity from directory objects.
        /// </summary>
        /// <param name="page"></param>
        /// <param name="groupGraphUsers"></param>
        /// <param name="groupUsers"></param>
        /// <param name="accessToken"></param>
        /// <returns></returns>

        private static List<User> GenerateNestedGraphUserCollection(IList<DirectoryObject> page, List<User> groupGraphUsers, List<GroupUser> groupUsers, string accessToken)
        {
            // Create a collection of Microsoft.Graph.User type
            foreach (var usr in page)
            {

                if (usr != null)
                {
                    if (usr.GetType() == typeof(User))
                    {
                        groupGraphUsers.Add((User)usr);
                    }
                }
            }

            // Get groups within the group and users in that group
            List<Group> unifiedGroupGraphGroups = new List<Group>();
            GenerateGraphGroupCollection(page, unifiedGroupGraphGroups);
            foreach (Group unifiedGroupGraphGroup in unifiedGroupGraphGroups)
            {
                var grp = GetGroup(unifiedGroupGraphGroup.Id, accessToken);
                groupUsers.AddRange(GetGroupMembers(grp, accessToken));
            }

            return groupGraphUsers;
        }

        /// <summary>
        /// Helper method. Generates a collection of Microsoft.Graph.Group entity from directory objects.
        /// </summary>
        /// <param name="page"></param>
        /// <param name="groupGraphGroups"></param>
        /// <returns></returns>
        private static List<Group> GenerateGraphGroupCollection(IList<DirectoryObject> page, List<Group> groupGraphGroups)
        {
            // Create a collection of Microsoft.Graph.Group type
            foreach (var grp in page)
            {

                if (grp != null)
                {
                    if (grp.GetType() == typeof(Group))
                    {
                        groupGraphGroups.Add((Group)grp);
                    }
                }
            }

            return groupGraphGroups;
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
