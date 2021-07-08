using System;
using System.Net.Http;
using System.Net.Http.Headers;

namespace PnP.Framework.Graph
{
    public static class GraphHttpClient
    {
        public const string GraphBaseUrl = "https://graph.microsoft.com";
        [Obsolete("Use GraphHttpClient.GetGraphEndPointUrl(AzureEnvironment) instead.")]
        public const string MicrosoftGraphV1BaseUri = GraphBaseUrl + "/v1.0/";
        [Obsolete("Use GraphHttpClient.GetGraphEndPointUrl(AzureEnvironment, true) instead.")]
        public const string MicrosoftGraphBetaBaseUri = GraphBaseUrl + "/beta/";

        /// <summary>
        /// Returns the graph endpoint URL based upon the Azure Environment
        /// </summary>
        /// <param name="azureEnvironment"></param>
        /// <param name="beta"></param>
        /// <returns></returns>
        public static string GetGraphEndPointUrl(AzureEnvironment azureEnvironment = AzureEnvironment.Production, bool beta = false)
        {
            var endPoint = AuthenticationManager.GetGraphEndPoint(azureEnvironment);
            return $"https://{endPoint}/{(beta ? "beta/" : "v1.0/")}";
        }


        /// <summary>
        /// This helper method makes an HTTP GET request and returns the result as a String
        /// </summary>
        /// <param name="requestUrl">The URL of the request</param>
        /// <param name="accessToken">The AccessToken to use for the request</param>
        /// <returns>The string value of the result</returns>
        public static string MakeGetRequestForString(string requestUrl, string accessToken = null)
        {
            return (MakeHttpRequest<String>("GET",
                requestUrl,
                accessToken: accessToken,
                resultPredicate: r => r.Content.ReadAsStringAsync().Result));
        }

        /// <summary>
        /// This helper method makes an HTTP GET request and returns the result as a String
        /// </summary>
        /// <param name="requestUrl">The URL of the request</param>
        /// <param name="accept">The accept header for the response</param>
        /// <param name="accessToken">The AccessToken to use for the request</param>
        /// <returns>The Stream  of the result</returns>
        public static System.IO.Stream MakeGetRequestForStream(string requestUrl, string accept, string accessToken = null)
        {
            return (MakeHttpRequest<System.IO.Stream>("GET",
                requestUrl,
                accessToken: accessToken,
                resultPredicate: r => r.Content.ReadAsStreamAsync().Result));
        }

        /// <summary>
        /// This helper method makes an HTTP POST request without a response
        /// </summary>
        /// <param name="requestUrl">The URL of the request</param>
        /// <param name="content">The content of the request</param>
        /// <param name="contentType">The content/type of the request</param>
        /// <param name="accessToken">The AccessToken to use for the request</param>
        public static void MakePostRequest(string requestUrl, object content = null, string contentType = null, string accessToken = null)
        {
            MakeHttpRequest<string>("POST",
                requestUrl,
                content: content,
                contentType: contentType,
                accessToken: accessToken);
        }

        /// <summary>
        /// This helper method makes an HTTP POST request and returns the result as a String
        /// </summary>
        /// <param name="requestUrl">The URL of the request</param>
        /// <param name="content">The content of the request</param>
        /// <param name="contentType">The content/type of the request</param>
        /// <param name="accessToken">The AccessToken to use for the request</param>
        /// <returns>The String value of the result</returns>
        public static string MakePostRequestForString(string requestUrl, object content = null, string contentType = null, string accessToken = null)
        {
            return (MakeHttpRequest<string>("POST",
                requestUrl,
                content: content,
                contentType: contentType,
                accessToken: accessToken,
                resultPredicate: r => r.Content.ReadAsStringAsync().Result));
        }

        /// <summary>
        /// This helper method makes an HTTP PUT request without a response
        /// </summary>
        /// <param name="requestUrl">The URL of the request</param>
        /// <param name="content">The content of the request</param>
        /// <param name="contentType">The content/type of the request</param>
        /// <param name="accessToken">The AccessToken to use for the request</param>
        public static void MakePutRequest(string requestUrl, object content = null, string contentType = null, string accessToken = null)
        {
            MakeHttpRequest<string>("PUT",
                requestUrl,
                content: content,
                contentType: contentType,
                accessToken: accessToken);
        }

        /// <summary>
        /// This helper method makes an HTTP PATCH request and returns the result as a String
        /// </summary>
        /// <param name="requestUrl">The URL of the request</param>
        /// <param name="content">The content of the request</param>
        /// <param name="contentType">The content/type of the request</param>
        /// <param name="accessToken">The AccessToken to use for the request</param>
        /// <returns>The String value of the result</returns>
        public static string MakePatchRequestForString(string requestUrl, object content = null, string contentType = null, string accessToken = null)
        {
            return (MakeHttpRequest<string>("PATCH",
                requestUrl,
                content: content,
                contentType: contentType,
                accessToken: accessToken,
                resultPredicate: r => r.Content.ReadAsStringAsync().Result));
        }

        /// <summary>
        /// This helper method makes an HTTP DELETE request
        /// </summary>
        /// <param name="requestUrl">The URL of the request</param>
        /// <param name="accessToken">The AccessToken to use for the request</param>
        /// <returns>The String value of the result</returns>
        public static void MakeDeleteRequest(string requestUrl, string accessToken = null)
        {
            MakeHttpRequest<string>("DELETE", requestUrl, accessToken: accessToken);
        }

        /// <summary>
        /// This helper method makes an HTTP request and eventually returns a result
        /// </summary>
        /// <param name="httpMethod">The HTTP method for the request</param>
        /// <param name="requestUrl">The URL of the request</param>
        /// <param name="accept">The content type of the accepted response</param>
        /// <param name="content">The content of the request</param>
        /// <param name="contentType">The content  type of the request</param>
        /// <param name="resultPredicate">The predicate to retrieve the result, if any</param>
        /// <param name="accessToken">The AccessToken to use for the request</param>
        /// <typeparam name="TResult">The type of the result, if any</typeparam>
        /// <returns>The value of the result, if any</returns>
        private static TResult MakeHttpRequest<TResult>(
            string httpMethod,
            string requestUrl,
            string accept = null,
            object content = null,
            string contentType = null,
            string accessToken = null,
            Func<HttpResponseMessage, TResult> resultPredicate = null)
        {
            HttpResponseHeaders responseHeaders;

            return PnP.Framework.Utilities.HttpHelper.MakeHttpRequest(
                httpMethod,
                requestUrl,
                out responseHeaders,
                accessToken,
                accept,
                content,
                contentType,
                resultPredicate: resultPredicate);

        }
    }
}
