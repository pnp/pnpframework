using Microsoft.Graph;
using Newtonsoft.Json;
using PnP.Framework.Diagnostics;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Threading.Tasks;
using System.Web;

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
        /// <returns>List with User objects</returns>
        public static Model.User GetUser(string accessToken, Guid userId, string[] selectProperties = null, int startIndex = 0, int? endIndex = 999, int retryCount = 10, int delay = 500, bool useBetaEndPoint = false)
        {
            return ListUsers(accessToken, $"id eq '{userId}'", null, selectProperties, startIndex, endIndex, retryCount, delay, useBetaEndPoint: useBetaEndPoint).FirstOrDefault();
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
        /// <returns>User object</returns>
        public static Model.User GetUser(string accessToken, string userPrincipalName, string[] selectProperties = null, int startIndex = 0, int? endIndex = 999, int retryCount = 10, int delay = 500, bool useBetaEndPoint = false)
        {
            return ListUsers(accessToken, $"userPrincipalName eq '{userPrincipalName}'", null, selectProperties, startIndex, endIndex, retryCount, delay, useBetaEndPoint: useBetaEndPoint).FirstOrDefault();
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
        /// <returns>List with User objects</returns>
        public static List<Model.User> ListUsers(string accessToken, string[] additionalProperties = null, int startIndex = 0, int? endIndex = 999, int retryCount = 10, int delay = 500, bool useBetaEndPoint = false)
        {
            return ListUsers(accessToken, null, null, additionalProperties, startIndex, endIndex, retryCount, delay, useBetaEndPoint: useBetaEndPoint);
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
        /// <returns>List with User objects</returns>
        public static List<Model.User> ListUsers(string accessToken, string filter, string orderby, string[] selectProperties = null, int startIndex = 0, int? endIndex = 999, int retryCount = 10, int delay = 500, bool useBetaEndPoint = false)
        {
            if (String.IsNullOrEmpty(accessToken))
            {
                throw new ArgumentNullException(nameof(accessToken));
            }
            // Rewrite AdditionalProperties to Additional Data
            var propertiesToSelect = new List<string> { "BusinessPhones", "DisplayName", "GivenName", "JobTitle", "Mail", "MobilePhone", "OfficeLocation", "PreferredLanguage", "Surname", "UserPrincipalName", "Id", "AccountEnabled" };
            
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

            List<Model.User> result = null;
            try
            {
                // Use a synchronous model to invoke the asynchronous process
                result = Task.Run(async () =>
                {
                    List<Model.User> users = new List<Model.User>();

                    var graphClient = GraphUtility.CreateGraphClient(accessToken, retryCount, delay, useBetaEndPoint: useBetaEndPoint);

                    IGraphServiceUsersCollectionPage pagedUsers;

                    // Retrieve the first batch of users. 999 is the maximum amount of users that Graph allows to be trieved in 1 go. Use maximum size batches to lessen the chance of throttling when retrieving larger amounts of users.
                    pagedUsers = await graphClient.Users.Request()
                                                        .Select(string.Join(",", propertiesToSelect))
                                                        .Filter(filter)
                                                        .OrderBy(orderby)
                                                        .Top(!endIndex.HasValue ? 999 : endIndex.Value >= 999 ? 999 : endIndex.Value)
                                                        .GetAsync();

                    int pageCount = 0;
                    int currentIndex = 0;

                    while (true)
                    {
                        pageCount++;

                        foreach (var u in pagedUsers)
                        {
                            currentIndex++;

                            if(endIndex.HasValue && endIndex.Value < currentIndex)
                            {
                                break;
                            }

                            if (currentIndex >= startIndex)
                            {
                                var user = new Model.User
                                {
                                    Id = Guid.TryParse(u.Id, out Guid idGuid) ? (Guid?)idGuid : null,
                                    DisplayName = u.DisplayName,
                                    GivenName = u.GivenName,
                                    JobTitle = u.JobTitle,
                                    MobilePhone = u.MobilePhone,
                                    OfficeLocation = u.OfficeLocation,
                                    PreferredLanguage = u.PreferredLanguage,
                                    Surname = u.Surname,
                                    UserPrincipalName = u.UserPrincipalName,
                                    BusinessPhones = u.BusinessPhones,
                                    AdditionalProperties = u.AdditionalData,
                                    Mail = u.Mail,
                                    AccountEnabled = u.AccountEnabled,
                                };

                                // If additional properties have been provided, ensure their output gets added to the AdditionalProperties dictonary of the output
                                if (selectProperties != null)
                                {
                                    // Ensure we have the AdditionalProperties dictionary available to fill, if necessary
                                    if(user.AdditionalProperties == null)
                                    {
                                        user.AdditionalProperties = new Dictionary<
                                        string, object>();
                                    }

                                    foreach (var selectProperty in selectProperties)
                                    {
                                        // Ensure the requested property has been returned in the response
                                        var property = u.GetType().GetProperty(selectProperty, BindingFlags.IgnoreCase |  BindingFlags.Public | BindingFlags.Instance);
                                        if (property != null)
                                        {
                                            // First check if we have the property natively on the User model
                                            var userProperty = user.GetType().GetProperty(selectProperty, BindingFlags.IgnoreCase |  BindingFlags.Public | BindingFlags.Instance);
                                            if(userProperty != null)
                                            {
                                                // Set the property on the User model
                                                userProperty.SetValue(user, property.GetValue(u), null);
                                            }
                                            else
                                            {
                                                // Property does not exist on the User model, add the property to the AdditionalProperties dictionary
                                                user.AdditionalProperties.Add(selectProperty, property.GetValue(u));
                                            }
                                        }
                                    }
                                }

                                users.Add(user);
                            }
                        }

                        if (pagedUsers.NextPageRequest != null && (!endIndex.HasValue || currentIndex < endIndex.Value))
                        {
                            // Retrieve the next batch of users. The possible oData instructions such as select and filter are already incorporated in the nextLink provided by Graph and thus do not need to be specified again.
                            pagedUsers = await pagedUsers.NextPageRequest.GetAsync();
                        }
                        else
                        {
                            break;
                        }
                    }

                    return users;
                }).GetAwaiter().GetResult();
            }
            catch (ServiceException ex)
            {
                Log.Error(Constants.LOGGING_SOURCE, CoreResources.GraphExtensions_ErrorOccured, ex.Error.Message);
                throw;
            }
            return result;
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
        /// <param name="azureEnvironment">Defines the Azure Cloud Deployment. This is used to determine the MS Graph EndPoint to call which differs per Azure Cloud deployments.</param>
        /// <param name="useBetaEndPoint">Indicates if the v1.0 (false) or beta (true) endpoint should be used at Microsoft Graph to query for the data</param>
        /// <returns>List with User objects</returns>
        public static Model.UserDelta ListUserDelta(string accessToken, string deltaToken, string filter, string orderby, string[] selectProperties = null, int startIndex = 0, int? endIndex = 999, int retryCount = 10, int delay = 500, AzureEnvironment azureEnvironment = AzureEnvironment.Production, bool useBetaEndPoint = false)
        {
            if (String.IsNullOrEmpty(accessToken))
            {
                throw new ArgumentNullException(nameof(accessToken));
            }

            Model.UserDelta userDelta = new Model.UserDelta
            {
                Users = new List<Model.User>()
            };
            try
            {
                // GET https://graph.microsoft.com/v1.0/users/delta
                string getUserDeltaUrl = $"{GraphHttpClient.GetGraphEndPointUrl(azureEnvironment, useBetaEndPoint)}users/delta?";

                if (selectProperties != null)
                {
                    getUserDeltaUrl += $"$select={string.Join(",", selectProperties)}&";
                }
                if (!string.IsNullOrEmpty(filter))
                {
                    getUserDeltaUrl += $"$filter={filter}&";
                }
                if (!string.IsNullOrEmpty(deltaToken))
                {
                    getUserDeltaUrl += $"$deltatoken={deltaToken}&";
                }
                if (!string.IsNullOrEmpty(orderby))
                {
                    getUserDeltaUrl += $"$orderby={orderby}&";
                }

                getUserDeltaUrl = getUserDeltaUrl.TrimEnd('&').TrimEnd('?');

                int currentIndex = 0;

                while (true)
                {
                    var response = GraphHttpClient.MakeGetRequestForString(
                        requestUrl: getUserDeltaUrl,
                        accessToken: accessToken);

                    var userDeltaResponse = JsonConvert.DeserializeObject<Model.UserDelta>(response);

                    if (!string.IsNullOrEmpty(userDeltaResponse.DeltaToken))
                    {
                        userDelta.DeltaToken = HttpUtility.ParseQueryString(new Uri(userDeltaResponse.DeltaToken).Query).Get("$deltatoken");
                    }

                    foreach (var user in userDeltaResponse.Users)
                    {
                        currentIndex++;

                        if(endIndex.HasValue && endIndex.Value < currentIndex)
                        {
                            break;
                        }

                        if (currentIndex >= startIndex && (!endIndex.HasValue || currentIndex <= endIndex.Value))
                        {
                            userDelta.Users.Add(user);
                        }
                    }

                    if (userDeltaResponse.NextLink != null && (!endIndex.HasValue || currentIndex < endIndex.Value))
                    {
                        getUserDeltaUrl = userDeltaResponse.NextLink;
                    }
                    else
                    {
                        break;
                    }
                }
            }
            catch (ServiceException ex)
            {
                Log.Error(Constants.LOGGING_SOURCE, CoreResources.GraphExtensions_ErrorOccured, ex.Error.Message);
                throw;
            }
            return userDelta;
        }
    }
}
