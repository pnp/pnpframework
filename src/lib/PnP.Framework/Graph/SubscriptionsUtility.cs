using PnP.Framework.Diagnostics;
using PnP.Framework.Utilities;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http;
using System.Text.Json;
using System.Text.Json.Nodes;

namespace PnP.Framework.Graph
{
    /// <summary>
    /// Class that deals with Microsoft Graph Subscriptions
    /// </summary>
    public static class SubscriptionsUtility
    {
        /// <summary>
        /// Returns the subscription with the provided subscriptionId from Microsoft Graph
        /// </summary>
        /// <param name="accessToken">The OAuth 2.0 Access Token to use for invoking the Microsoft Graph</param>
        /// <param name="subscriptionId">The unique identifier of the subscription to return from Microsoft Graph</param>        
        /// <param name="startIndex">First item in the results returned by Microsoft Graph to return</param>
        /// <param name="endIndex">Last item in the results returned by Microsoft Graph to return</param>
        /// <param name="retryCount">Number of times to retry the request in case of throttling</param>
        /// <param name="delay">Milliseconds to wait before retrying the request. The delay will be increased (doubled) every retry.</param>
        /// <param name="azureEnvironment">Defines the Azure Cloud Deployment. This is used to determine the MS Graph EndPoint to call which differs per Azure Cloud deployments. Defaults to Production (graph.microsoft.com).</param>
        /// <returns>Subscription object</returns>
        public static Model.Subscription GetSubscription(string accessToken, Guid subscriptionId, int startIndex = 0, int endIndex = 999, int retryCount = 10, int delay = 500, AzureEnvironment azureEnvironment = AzureEnvironment.Production)
        {
            try
            {
                var requestUrl = $"{GraphHttpClient.GetGraphEndPointUrl(azureEnvironment)}subscriptions/{subscriptionId}";

                var responseAsString = HttpHelper.MakeGetRequestForString(requestUrl, accessToken, retryCount: retryCount, delay: delay);
                var jsonNode = JsonNode.Parse(responseAsString);
                var subscription = jsonNode["value"];

                var subscriptionModel = subscription.Deserialize<Model.Subscription>();
                return subscriptionModel;
            }
            catch (ApplicationException ex)
            {
                Log.Error(Constants.LOGGING_SOURCE, ex.Message);
                throw;
            }
        }

        /// <summary>
        /// Returns all the active Microsoft Graph subscriptions
        /// </summary>
        /// <param name="accessToken">The OAuth 2.0 Access Token to use for invoking the Microsoft Graph</param>
        /// <param name="startIndex">First item in the results returned by Microsoft Graph to return</param>
        /// <param name="endIndex">Last item in the results returned by Microsoft Graph to return</param>
        /// <param name="retryCount">Number of times to retry the request in case of throttling</param>
        /// <param name="delay">Milliseconds to wait before retrying the request. The delay will be increased (doubled) every retry.</param>
        /// <param name="azureEnvironment">Defines the Azure Cloud Deployment. This is used to determine the MS Graph EndPoint to call which differs per Azure Cloud deployments. Defaults to Production (graph.microsoft.com).</param>
        /// <returns>List with Subscription objects</returns>
        public static List<Model.Subscription> ListSubscriptions(string accessToken, int startIndex = 0, int endIndex = 999, int retryCount = 10, int delay = 500, AzureEnvironment azureEnvironment = AzureEnvironment.Production)
        {
            if (String.IsNullOrEmpty(accessToken))
            {
                throw new ArgumentNullException(nameof(accessToken));
            }

            List<Model.Subscription> result = new();
            try
            {
                var requestUrl = $"{GraphHttpClient.GetGraphEndPointUrl(azureEnvironment)}subscriptions";

                int currentIndex = 0;

                do
                {
                    var responseAsString = HttpHelper.MakeGetRequestForString(requestUrl, accessToken, retryCount: retryCount, delay: delay);
                    var jsonNode = JsonNode.Parse(responseAsString);
                    var subscriptionListString = jsonNode["value"];

                    var subscriptionsPage = subscriptionListString.Deserialize<Model.Subscription[]>();

                    var startIndexForPage = Math.Min(0, startIndex - (currentIndex));
                    var endIndexForPage = Math.Min(subscriptionsPage.Length - 1, endIndex - currentIndex);

                    // Todo - test index values, might be off by 1
                    result.AddRange(subscriptionsPage.Skip(startIndexForPage).Take(startIndexForPage - endIndexForPage));

                    currentIndex += subscriptionsPage.Length;

                    if (currentIndex >= endIndex)
                    {
                        break;
                    }

                    requestUrl = jsonNode["@odata.nextLink"]?.ToString();

                } while (!string.IsNullOrEmpty(requestUrl));
                
            }
            catch (HttpRequestException ex)
            {
                Log.Error(Constants.LOGGING_SOURCE, CoreResources.GraphExtensions_ErrorOccured, ex.Message);
                throw;
            }
            return result;
        }

        /// <summary>
        /// Creates a new Microsoft Graph Subscription
        /// </summary>
        /// <param name="changeType">The event(s) the subscription should trigger on</param>
        /// <param name="notificationUrl">The URL that should be called when an event matching this subscription occurs</param>
        /// <param name="resource">The resource to monitor for changes. See https://docs.microsoft.com/graph/api/subscription-post-subscriptions#permissions for the list with supported options.</param>
        /// <param name="expirationDateTime">The datetime defining how long this subscription should stay alive before which it needs to get extended to stay alive. See https://docs.microsoft.com/graph/api/resources/subscription#maximum-length-of-subscription-per-resource-type for the supported maximum lifetime of the subscriber endpoints.</param>
        /// <param name="clientState">Specifies the value of the clientState property sent by the service in each notification. The maximum length is 128 characters. The client can check that the notification came from the service by comparing the value of the clientState property sent with the subscription with the value of the clientState property received with each notification.</param>
        /// <param name="latestSupportedTlsVersion">Specifies the latest version of Transport Layer Security (TLS) that the notification endpoint, specified by <paramref name="notificationUrl"/>, supports. If not provided, TLS 1.2 will be assumed.</param>
        /// <param name="accessToken">The OAuth 2.0 Access Token to use for invoking the Microsoft Graph</param>
        /// <param name="retryCount">Number of times to retry the request in case of throttling</param>
        /// <param name="delay">Milliseconds to wait before retrying the request. The delay will be increased (doubled) every retry</param>
        /// <param name="azureEnvironment">Defines the Azure Cloud Deployment. This is used to determine the MS Graph EndPoint to call which differs per Azure Cloud deployments. Defaults to Production (graph.microsoft.com).</param>
        /// <returns>The just created Microsoft Graph subscription</returns>
        public static Model.Subscription CreateSubscription(Enums.GraphSubscriptionChangeType changeType, string notificationUrl, string resource, DateTimeOffset expirationDateTime, string clientState,
                                                            string accessToken, Enums.GraphSubscriptionTlsVersion latestSupportedTlsVersion = Enums.GraphSubscriptionTlsVersion.v1_2, int retryCount = 10, int delay = 500, AzureEnvironment azureEnvironment = AzureEnvironment.Production)
        {
            if (String.IsNullOrEmpty(notificationUrl))
            {
                throw new ArgumentNullException(nameof(notificationUrl));
            }

            if (String.IsNullOrEmpty(resource))
            {
                throw new ArgumentNullException(nameof(resource));
            }

            try
            {
                var requestUrl = $"{GraphHttpClient.GetGraphEndPointUrl(azureEnvironment)}subscriptions";
                var newSubscription = new
                {
                        ChangeType = changeType.ToString().Replace(" ", ""),
                        NotificationUrl = notificationUrl,
                        Resource = resource,
                        ExpirationDateTime = expirationDateTime,
                        ClientState = clientState
                    };
                var stringContent = JsonSerializer.Serialize(newSubscription);
                var content = new StringContent(stringContent);

                var responseAsString = HttpHelper.MakePostRequestForString(requestUrl, content, "application/json", accessToken, retryCount: retryCount, delay: delay);

                // Todo - check that the returned data does actually deserialise correctly
                var model = JsonSerializer.Deserialize<Model.Subscription>(responseAsString);
                return model;
                    }
            catch (HttpResponseException ex)
            {
                Log.Error(Constants.LOGGING_SOURCE, CoreResources.GraphExtensions_ErrorOccured, ex.Message);
                throw;
            }
        }

        /// <summary>
        /// Updates an existing Microsoft Graph Subscription
        /// </summary>
        /// <param name="subscriptionId">Unique identifier of the Microsoft Graph subscription</param>
        /// <param name="expirationDateTime">The datetime defining how long this subscription should stay alive before which it needs to get extended to stay alive. See https://docs.microsoft.com/graph/api/resources/subscription#maximum-length-of-subscription-per-resource-type for the supported maximum lifetime of the subscriber endpoints.</param>
        /// <param name="accessToken">The OAuth 2.0 Access Token to use for invoking the Microsoft Graph</param>
        /// <param name="retryCount">Number of times to retry the request in case of throttling</param>
        /// <param name="delay">Milliseconds to wait before retrying the request. The delay will be increased (doubled) every retry</param>
        /// <param name="azureEnvironment">Defines the Azure Cloud Deployment. This is used to determine the MS Graph EndPoint to call which differs per Azure Cloud deployments. Defaults to Production (graph.microsoft.com).</param>
        /// <returns>The just updated Microsoft Graph subscription</returns>
        public static Model.Subscription UpdateSubscription(string subscriptionId, DateTimeOffset expirationDateTime, string accessToken, int retryCount = 10, int delay = 500, AzureEnvironment azureEnvironment = AzureEnvironment.Production)
        {
            if (String.IsNullOrEmpty(subscriptionId))
            {
                throw new ArgumentNullException(nameof(subscriptionId));
            }

            Model.Subscription result = null;

            try
            {
                var requestUrl = $"{GraphHttpClient.GetGraphEndPointUrl(azureEnvironment)}subscriptions/{subscriptionId}";
                var updatedSubscription = new Model.Subscription
                {
                    ExpirationDateTime = expirationDateTime
                };
                var contentString = JsonSerializer.Serialize(updatedSubscription);
                var content = new StringContent(contentString);

                var responseAsString = HttpHelper.MakePatchRequestForString(requestUrl, content, "application/json", accessToken, retryCount: retryCount, delay: delay);

                // Todo - check that the returned data does actually deserialise correctly
                var model = JsonSerializer.Deserialize<Model.Subscription>(responseAsString);
                return model;
            }
            catch (HttpRequestException ex)
            {
                Log.Error(Constants.LOGGING_SOURCE, CoreResources.GraphExtensions_ErrorOccured, ex.Message);
                throw;
            }
        }

        /// <summary>
        /// Deletes an existing Microsoft Graph Subscription
        /// </summary>
        /// <param name="subscriptionId">Unique identifier of the Microsoft Graph subscription</param>
        /// <param name="accessToken">The OAuth 2.0 Access Token to use for invoking the Microsoft Graph</param>
        /// <param name="retryCount">Number of times to retry the request in case of throttling</param>
        /// <param name="delay">Milliseconds to wait before retrying the request. The delay will be increased (doubled) every retry</param>
        public static void DeleteSubscription(string subscriptionId,
                                              string accessToken, int retryCount = 10, int delay = 500)
        {
            if (String.IsNullOrEmpty(subscriptionId))
            {
                throw new ArgumentNullException(nameof(subscriptionId));
            }

            try
            {
                var requestUrl = $"{GraphHttpClient.GetGraphEndPointUrl()}subscriptions/{subscriptionId}";

                HttpHelper.MakeDeleteRequest(requestUrl, accessToken, retryCount: retryCount, delay: delay);
            }
            catch (HttpRequestException ex)
            {
                Log.Error(Constants.LOGGING_SOURCE, CoreResources.GraphExtensions_ErrorOccured, ex.Message);
                throw;
            }
        }
    }
}
