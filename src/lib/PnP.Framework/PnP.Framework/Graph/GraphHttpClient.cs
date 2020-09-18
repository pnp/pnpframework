using System;
using System.Net.Http;
using System.Net.Http.Headers;

namespace PnP.Framework.Graph
{
    public static class GraphHttpClient
    {
        public const String GraphBaseUrl = "https://graph.microsoft.com";
        public const String MicrosoftGraphV1BaseUri = GraphBaseUrl + "/v1.0/";
        public const String MicrosoftGraphBetaBaseUri = GraphBaseUrl + "/beta/";

        /// <summary>
        /// This helper method makes an HTTP GET request and returns the result as a String
        /// </summary>
        /// <param name="requestUrl">The URL of the request</param>
        /// <param name="accessToken">The AccessToken to use for the request</param>
        /// <returns>The String value of the result</returns>
        public static String MakeGetRequestForString(String requestUrl,
            String accessToken = null)
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
        public static System.IO.Stream MakeGetRequestForStream(String requestUrl,
            String accept, String accessToken = null)
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
        public static void MakePostRequest(String requestUrl,
            Object content = null,
            String contentType = null,
            String accessToken = null)
        {
            MakeHttpRequest<String>("POST",
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
        public static String MakePostRequestForString(String requestUrl,
            Object content = null,
            String contentType = null,
            String accessToken = null)
        {
            return (MakeHttpRequest<String>("POST",
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
        public static void MakePutRequest(String requestUrl,
            Object content = null,
            String contentType = null,
            String accessToken = null)
        {
            MakeHttpRequest<String>("PUT",
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
        public static String MakePatchRequestForString(String requestUrl,
            Object content = null,
            String contentType = null,
            String accessToken = null)
        {
            return (MakeHttpRequest<String>("PATCH",
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
        public static void MakeDeleteRequest(String requestUrl,
            String accessToken = null)
        {
            MakeHttpRequest<String>("DELETE", requestUrl, accessToken: accessToken);
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
            String httpMethod,
            String requestUrl,
            String accept = null,
            Object content = null,
            String contentType = null,
            String accessToken = null,
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
