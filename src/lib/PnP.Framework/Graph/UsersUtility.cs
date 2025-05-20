using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using PnP.Framework.Diagnostics;
using PnP.Framework.Utilities;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.Json.Nodes;
using System.Text.Json;

namespace PnP.Framework.Graph
{
    /// <summary>
    /// Provides access to user operations in Microsoft Graph
    /// </summary>
    public static class UsersUtility
    {
        /// <summary>
        /// Returns the user with the provided userId from Azure Active Directory
        /// </summary>
        /// <param name="accessToken">The OAuth 2.0 Access Token to use for invoking the Microsoft Graph</param>
        /// <param name="userId">The unique identifier of the user in Azure Active Directory to return</param>    
        /// <param name="selectProperties">Allows providing the names of properties to return regarding the users. If not provided, the standard properties will be returned.</param>
        /// <param name="startIndex">First item in the results returned by Microsoft Graph to return</param>
        /// <param name="endIndex">Last item in the results returned by Microsoft Graph to return. Provide NULL to return all results that exist.</param>
        /// <param name="retryCount">Number of times to retry the request in case of throttling</param>
        /// <param name="delay">Milliseconds to wait before retrying the request. The delay will be increased (doubled) every retry.</param>
        /// <param name="useBetaEndPoint">Indicates if the v1.0 (false) or beta (true) endpoint should be used at Microsoft Graph to query for the data</param>
        /// <param name="ignoreDefaultProperties">If set to true, only the properties provided through selectProperties will be loaded. The default properties will not be. Optional. Default is that the default properties will always be retrieved.</param>
        /// <param name="azureEnvironment">The type of environment to connect to</param>
        /// <returns>List with User objects</returns>
        public static Model.User GetUser(string accessToken, Guid userId, string[] selectProperties = null, int startIndex = 0, int? endIndex = 999, int retryCount = 10, int delay = 500, bool useBetaEndPoint = false, bool ignoreDefaultProperties = false, AzureEnvironment azureEnvironment = AzureEnvironment.Production)
        {
            return ListUsers(accessToken, $"id eq '{userId}'", null, selectProperties, startIndex, endIndex, retryCount, delay, ignoreDefaultProperties: ignoreDefaultProperties, useBetaEndPoint: useBetaEndPoint, azureEnvironment: azureEnvironment).FirstOrDefault();
        }

        /// <summary>
        /// Returns the user with the provided <paramref name="userPrincipalName"/> from Azure Active Directory
        /// </summary>
        /// <param name="accessToken">The OAuth 2.0 Access Token to use for invoking the Microsoft Graph</param>
        /// <param name="userPrincipalName">The User Principal Name of the user in Azure Active Directory to return</param>
        /// <param name="selectProperties">Allows providing the names of properties to return regarding the users. If not provided, the standard properties will be returned.</param>
        /// <param name="startIndex">First item in the results returned by Microsoft Graph to return</param>
        /// <param name="endIndex">Last item in the results returned by Microsoft Graph to return. Provide NULL to return all results that exist.</param>
        /// <param name="retryCount">Number of times to retry the request in case of throttling</param>
        /// <param name="delay">Milliseconds to wait before retrying the request. The delay will be increased (doubled) every retry.</param>
        /// <param name="useBetaEndPoint">Indicates if the v1.0 (false) or beta (true) endpoint should be used at Microsoft Graph to query for the data</param>
        /// <param name="ignoreDefaultProperties">If set to true, only the properties provided through selectProperties will be loaded. The default properties will not be. Optional. Default is that the default properties will always be retrieved.</param>
        /// <returns>User object</returns>
        public static Model.User GetUser(string accessToken, string userPrincipalName, string[] selectProperties = null, int startIndex = 0, int? endIndex = 999, int retryCount = 10, int delay = 500, bool useBetaEndPoint = false, bool ignoreDefaultProperties = false, AzureEnvironment azureEnvironment = AzureEnvironment.Production)
        {
            return ListUsers(accessToken, $"userPrincipalName eq '{userPrincipalName}'", null, selectProperties, startIndex, endIndex, retryCount, delay, ignoreDefaultProperties: ignoreDefaultProperties, useBetaEndPoint: useBetaEndPoint, azureEnvironment: azureEnvironment).FirstOrDefault();
        }

        /// <summary>
        /// Returns all the Users in the current domain
        /// </summary>
        /// <param name="accessToken">The OAuth 2.0 Access Token to use for invoking the Microsoft Graph</param> 
        /// <param name="additionalProperties">Allows providing the names of additional properties to return regarding the users</param>
        /// <param name="startIndex">First item in the results returned by Microsoft Graph to return</param>
        /// <param name="endIndex">Last item in the results returned by Microsoft Graph to return. Provide NULL to return all results that exist.</param>
        /// <param name="retryCount">Number of times to retry the request in case of throttling</param>
        /// <param name="delay">Milliseconds to wait before retrying the request. The delay will be increased (doubled) every retry.</param>
        /// <param name="useBetaEndPoint">Indicates if the v1.0 (false) or beta (true) endpoint should be used at Microsoft Graph to query for the data</param>
        /// <param name="ignoreDefaultProperties">If set to true, only the properties provided through selectProperties will be loaded. The default properties will not be. Optional. Default is that the default properties will always be retrieved.</param>
        /// <param name="azureEnvironment">The type of environment to connect to</param>
        /// <returns>List with User objects</returns>
        public static List<Model.User> ListUsers(string accessToken, string[] additionalProperties = null, int startIndex = 0, int? endIndex = 999, int retryCount = 10, int delay = 500, bool useBetaEndPoint = false, bool ignoreDefaultProperties = false, AzureEnvironment azureEnvironment = AzureEnvironment.Production)
        {
            return ListUsers(accessToken, null, null, additionalProperties, startIndex, endIndex, retryCount, delay, ignoreDefaultProperties: ignoreDefaultProperties, useBetaEndPoint: useBetaEndPoint, azureEnvironment: azureEnvironment);
        }

        /// <summary>
        /// Returns all the Users in the current domain filtered out with a custom OData filter
        /// </summary>
        /// <param name="accessToken">The OAuth 2.0 Access Token to use for invoking the Microsoft Graph</param>
        /// <param name="filter">OData filter to apply to retrieval of the users from the Microsoft Graph</param>
        /// <param name="orderby">OData orderby instruction</param>
        /// <param name="selectProperties">Allows providing the names of properties to return regarding the users. If not provided, the standard properties will be returned.</param>
        /// <param name="startIndex">First item in the results returned by Microsoft Graph to return</param>
        /// <param name="endIndex">Last item in the results returned by Microsoft Graph to return. Provide NULL to return all results that exist.</param>
        /// <param name="retryCount">Number of times to retry the request in case of throttling</param>
        /// <param name="delay">Milliseconds to wait before retrying the request. The delay will be increased (doubled) every retry.</param>
        /// <param name="useBetaEndPoint">Indicates if the v1.0 (false) or beta (true) endpoint should be used at Microsoft Graph to query for the data</param>
        /// <param name="ignoreDefaultProperties">If set to true, only the properties provided through selectProperties will be loaded. The default properties will not be. Optional. Default is that the default properties will always be retrieved.</param>
        /// <returns>List with User objects</returns>
        public static List<Model.User> ListUsers(string accessToken, string filter, string orderby, string[] selectProperties = null, int startIndex = 0, int? endIndex = 999, int retryCount = 10, int delay = 500, bool useBetaEndPoint = false, bool ignoreDefaultProperties = false, AzureEnvironment azureEnvironment = AzureEnvironment.Production)
        {
            if (String.IsNullOrEmpty(accessToken))
            {
                throw new ArgumentNullException(nameof(accessToken));
            }
            // Rewrite AdditionalProperties to Additional Data
            var propertiesToSelect = ignoreDefaultProperties ? new List<string>() : new List<string> { "BusinessPhones", "DisplayName", "GivenName", "JobTitle", "Mail", "MobilePhone", "OfficeLocation", "PreferredLanguage", "Surname", "UserPrincipalName", "Id", "AccountEnabled" };
            
            selectProperties = selectProperties?.Select(p => p == "AdditionalProperties" ? "AdditionalData" : p).ToArray();
            
            if(selectProperties != null)
            {
                foreach(var property in selectProperties)
                {
                    if(!propertiesToSelect.Contains(property))
                    {
                        propertiesToSelect.Add(property);
                    }
                }
            }

            try
            {
                var requestUrl = $"{GraphHttpClient.GetGraphEndPointUrl(azureEnvironment)}users";
                var queryStringParams = new List<string>()
                {
                    $"$top={(!endIndex.HasValue ? 999 : endIndex.Value >= 999 ? 999 : endIndex.Value)}"
                };
                if (propertiesToSelect.Count > 0)
                {
                    queryStringParams.Add("$select=" + string.Join(",", propertiesToSelect));
                }
                if (!string.IsNullOrEmpty(filter))
                {
                    queryStringParams.Add($"$filter={filter}");
                }
                if (!string.IsNullOrEmpty(orderby))
                {
                    queryStringParams.Add($"orderby={orderby}");
                }
                requestUrl += $"?{string.Join("&", queryStringParams)}";
                IEnumerable<Model.User> users = GraphUtility.ReadPagedDataFromRequest<Model.User>(requestUrl, accessToken, retryCount: retryCount, delay: delay)
                                    .Skip(startIndex);
                if (endIndex.HasValue)
                {
                    users = users.Take(endIndex.Value - startIndex);
                }
                return users.ToList();

            }
            catch (HttpResponseException ex)
            {
                Log.Error(Constants.LOGGING_SOURCE, CoreResources.GraphExtensions_ErrorOccured, ex.Message);
                throw;
            }
        }

        /// <summary>
        /// Returns the users delta in the current domain filtered out with a custom OData filter. If no <paramref name="deltaToken"/> has been provided, all users will be returned with a deltatoken for a next run. If a <paramref name="deltaToken"/> has been provided, all users which were modified after the deltatoken has been generated will be returned.
        /// </summary>
        /// <param name="accessToken">The OAuth 2.0 Access Token to use for invoking the Microsoft Graph</param>
        /// <param name="deltaToken">DeltaToken to indicate requesting changes since this deltatoken has been created. Leave NULL to retrieve all users with a deltatoken to use for subsequent queries.</param>
        /// <param name="filter">OData filter to apply to retrieval of the users from the Microsoft Graph</param>
        /// <param name="orderby">OData orderby instruction</param>
        /// <param name="selectProperties">Allows providing the names of properties to return regarding the users. If not provided, the standard properties will be returned.</param>
        /// <param name="startIndex">First item in the results returned by Microsoft Graph to return</param>
        /// <param name="endIndex">Last item in the results returned by Microsoft Graph to return. Provide NULL to return all results that exist.</param>
        /// <param name="retryCount">Number of times to retry the request in case of throttling</param>
        /// <param name="delay">Milliseconds to wait before retrying the request. The delay will be increased (doubled) every retry.</param>
        /// <param name="useBetaEndPoint">Indicates if the v1.0 (false) or beta (true) endpoint should be used at Microsoft Graph to query for the data</param>
        /// <param name="ignoreDefaultProperties">If set to true, only the properties provided through selectProperties will be loaded. The default properties will not be. Optional. Default is that the default properties will always be retrieved.</param>
        /// <param name="azureEnvironment">The type of environment to connect to</param>
        /// <returns>List with User objects</returns>
        public static Model.UserDelta ListUserDelta(string accessToken, string deltaToken, string filter, string orderby, string[] selectProperties = null, int startIndex = 0, int? endIndex = 999, int retryCount = 10, int delay = 500, bool useBetaEndPoint = false, bool ignoreDefaultProperties = false, AzureEnvironment azureEnvironment = AzureEnvironment.Production)
        {
            if (String.IsNullOrEmpty(accessToken))
            {
                throw new ArgumentNullException(nameof(accessToken));
            }
            // Rewrite AdditionalProperties to Additional Data
            var propertiesToSelect = ignoreDefaultProperties ? new List<string>() : new List<string> { "BusinessPhones", "DisplayName", "GivenName", "JobTitle", "Mail", "MobilePhone", "OfficeLocation", "PreferredLanguage", "Surname", "UserPrincipalName", "Id", "AccountEnabled" };

            selectProperties = selectProperties?.Select(p => p == "AdditionalProperties" ? "AdditionalData" : p).ToArray();

            if (selectProperties != null)
            {
                foreach (var property in selectProperties)
                {
                    if (!propertiesToSelect.Contains(property))
                    {
                        propertiesToSelect.Add(property);
                    }
                }
            }
            List<Model.User> users = new List<Model.User>();

            var requestUrl = $"{GraphHttpClient.GetGraphEndPointUrl(azureEnvironment)}users";
            var queryStringParams = new List<string>()
                {
                    $"$top={(!endIndex.HasValue ? 999 : endIndex.Value >= 999 ? 999 : endIndex.Value)}",
                    $"$skiptoken={deltaToken}",
                };
            if (propertiesToSelect.Count > 0)
            {
                queryStringParams.Add("$select=" + string.Join(",", propertiesToSelect));
            }
            if (!string.IsNullOrEmpty(filter))
            {
                queryStringParams.Add($"$filter={filter}");
            }
            if (!string.IsNullOrEmpty(orderby))
            {
                queryStringParams.Add($"orderby={orderby}");
            }
            requestUrl += $"?{string.Join("&", queryStringParams)}";

            try
            {
                int currentIndex = 0;
                var usersDelta = new Model.UserDelta();
                usersDelta.Users = users;


                while (requestUrl != null)
                {
                    var responseData = HttpHelper.MakeGetRequestForString(requestUrl, accessToken, retryCount: retryCount, delay: delay);

                    var jsonNode = JsonNode.Parse(responseData);
                    JsonNode valueNode = jsonNode["value"];
                    var results = valueNode.Deserialize<Model.User[]>(GraphUtility.CaseInsensitiveJsonOptions);

                    foreach (var r in results)
                    {
                        currentIndex++;

                        if (endIndex.HasValue && endIndex.Value < currentIndex)
                        {
                            break;
                        }
                        if (currentIndex >= startIndex)
                        {
                            users.Add(r);
                        }
                    }

                    usersDelta.NextLink = jsonNode["@odata.nextLink"]?.ToString();
                    requestUrl = (endIndex.HasValue && endIndex.Value < currentIndex) ? null : usersDelta.NextLink;

                    var deltaLink = jsonNode["@odata.deltalink"]?.ToString();

                    if (string.IsNullOrWhiteSpace(deltaLink))
                    {
                        // Use a regular expression to fetch just the deltatoken part from the deltalink. The base of the URL will thereby be cut off. This is the only part we need to use it in a subsequent run.
                        var deltaLinkMatch = System.Text.RegularExpressions.Regex.Match(deltaLink, @"(?<=\$deltatoken=)(.*?)(?=$|&)", System.Text.RegularExpressions.RegexOptions.IgnoreCase);

                        if (deltaLinkMatch.Success && !string.IsNullOrWhiteSpace(deltaLinkMatch.Value))
                        {
                            // Successfully extracted the deltatoken part from the link, assign it to the return variable
                            usersDelta.DeltaToken = deltaLinkMatch.Value;
                        }
                    }
                }
                return usersDelta;
            }
            catch (HttpResponseException ex)
            {
                Log.Error(Constants.LOGGING_SOURCE, CoreResources.GraphExtensions_ErrorOccured, ex.Message);
                throw;
            }
        }

        /// <summary>
        /// Returns deleted Users in the current domain
        /// </summary>
        /// <param name="accessToken">The OAuth 2.0 Access Token to use for invoking the Microsoft Graph</param>
        /// <param name="selectProperties">Allows providing the names of properties to return regarding the users. If not provided, the standard properties will be returned.</param>
        /// <param name="retryCount">Number of times to retry the request in case of throttling</param>
        /// <param name="delay">Milliseconds to wait before retrying the request. The delay will be increased (doubled) every retry.</param>
        /// <param name="ignoreDefaultProperties">If set to true, only the properties provided through selectProperties will be loaded. The default properties will not be. Optional. Default is that the default properties will always be retrieved.</param>
        /// /// <param name="azureEnvironment">The type of environment to connect to</param>
        /// <returns>List with User objects</returns>
        public static List<Model.User> ListDeletedUsers(string accessToken, string[] selectProperties = null, int retryCount = 10, int delay = 500, bool ignoreDefaultProperties = false, AzureEnvironment azureEnvironment = AzureEnvironment.Production)
        {
            if (String.IsNullOrEmpty(accessToken))
            {
                throw new ArgumentNullException(nameof(accessToken));
            }
            // Rewrite AdditionalProperties to Additional Data
            var propertiesToSelect = ignoreDefaultProperties ? new List<string>() : new List<string> { "BusinessPhones", "DisplayName", "GivenName", "JobTitle", "Mail", "MobilePhone", "OfficeLocation", "PreferredLanguage", "Surname", "UserPrincipalName", "Id", "AccountEnabled", "DeletedDateTime" };
            
            selectProperties = selectProperties?.Select(p => p == "AdditionalProperties" ? "AdditionalData" : p).ToArray();
            
            if(selectProperties != null)
            {
                foreach(var property in selectProperties)
                {
                    if(!propertiesToSelect.Contains(property))
                    {
                        propertiesToSelect.Add(property);
                    }
                }
            }
            
            try
            {
                List<Model.User> users = new List<Model.User>();
                var requestUrl = $"{GraphHttpClient.GetGraphEndPointUrl(azureEnvironment)}directory/deleteditems/microsoft.graph.user";
                if (propertiesToSelect.Count > 0)
                {
                    requestUrl += $"?$select={string.Join(",", propertiesToSelect)}";
                } 

                var responseAsString = HttpHelper.MakeGetRequestForString(requestUrl, accessToken, retryCount: retryCount, delay: delay);

                var response = JToken.Parse(responseAsString);
                var deletedUsers = response["value"];

                foreach (var deletedUser in deletedUsers)
                {
                    var user = deletedUser.ToObject<Model.User>();
                    users.Add(user);
                }

                return users;
            }
            catch (ApplicationException ex)
            {
                Log.Error(Constants.LOGGING_SOURCE, ex.Message);
                throw;
            }
        }

        /// <summary>
        /// Retrieves a temporary access pass for the provided user
        /// </summary>
        /// <param name="accessToken">The OAuth 2.0 Access Token to use for invoking the Microsoft Graph</param>
        /// <param name="userId">Id or user principal name of the user to request the access pass for</param>
        /// <param name="startDateTime">Date and time at which this access pass should become valid. Optional. If not provided, it immediately become valid.</param>
        /// <param name="lifeTimeInMinutes">Durationin minutes during which this access pass will be valid. Optional. If not provided, the default configured in Azure Active Directory will be used.</param>
        /// <param name="isUsableOnce">Boolean indicating if the access pass can be used to only log in once or repetitively during the lifetime of the access pass. Optional. If not provided, the default configured in Azure Active Directory will be used.</param>
        /// <param name="retryCount">Number of times to retry the request in case of throttling. Optional.</param>
        /// <param name="delay">Milliseconds to wait before retrying the request. The delay will be increased (doubled) every retry. Optional.</param>
        /// <param name="azureEnvironment">The type of environment to connect to</param>
        /// <returns>A temporary access pass for the provided user or NULL if unable to create a temporary access pass</returns>
        public static Model.TemporaryAccessPassResponse RequestTemporaryAccessPass(string accessToken, string userId, DateTime? startDateTime = null, int? lifeTimeInMinutes = null, bool? isUsableOnce = null, int retryCount = 10, int delay = 500, AzureEnvironment azureEnvironment = AzureEnvironment.Production)
        {
            if (String.IsNullOrEmpty(accessToken))
            {
                throw new ArgumentNullException(nameof(accessToken));
            }
            if (String.IsNullOrEmpty(userId))
            {
                throw new ArgumentNullException(nameof(userId));
            }            

            // Build the request body for the access pass
            var temporaryAccessPassAuthenticationMethod = new Model.TemporaryAccessPassRequest
            {
                StartDateTime = startDateTime?.ToUniversalTime(),
                LifetimeInMinutes = lifeTimeInMinutes,
                IsUsableOnce = isUsableOnce
            };

            try
            {
                // Request the access pass
                var response = GraphHttpClient.MakePostRequestForString(
                    requestUrl: $"{GraphHttpClient.GetGraphEndPointUrl(azureEnvironment, beta: true)}users/{userId}/authentication/temporaryAccessPassMethods",
                    content: temporaryAccessPassAuthenticationMethod,
                    contentType: HttpHelper.JsonContentType,
                    accessToken: accessToken);

                // Parse and return the response
                var accessPassResponse = JsonConvert.DeserializeObject<Model.TemporaryAccessPassResponse>(response);
                return accessPassResponse;

            }
            catch (HttpResponseException ex)
            {
                Log.Error(Constants.LOGGING_SOURCE, CoreResources.GraphExtensions_ErrorOccured, ex.Message);
                throw;
            }
        }        
    }
}