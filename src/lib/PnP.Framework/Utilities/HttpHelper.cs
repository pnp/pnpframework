﻿using Microsoft.SharePoint.Client;
using Newtonsoft.Json;
using PnP.Framework.Http;
using PnP.Framework.Provisioning.ObjectHandlers.Utilities;
using System;
using System.Collections.Generic;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Threading.Tasks;

namespace PnP.Framework.Utilities
{
    /// <summary>
    /// Static class full of helper methods to make HTTP requests
    /// </summary>
    public static class HttpHelper
    {
        public const string JsonContentType = "application/json";

        /// <summary>
        /// This helper method makes an HTTP GET request and returns the result as a String
        /// </summary>
        /// <param name="requestUrl">The URL of the request</param>
        /// <param name="accept">The value for the accept header in the request</param>
        /// <param name="accessToken">The OAuth 2.0 Access Token for the request, if authorization is required</param>
        /// <param name="requestHeaders">A collection of any custom request headers</param>
        /// <param name="cookies">Any request cookies values</param>
        /// <param name="retryCount">Number of times to retry the request</param>
        /// <param name="delay">Milliseconds to wait before retrying the request. The delay will be increased (doubled) every retry</param>
        /// <param name="userAgent">UserAgent string value to insert for this request. You can define this value in your app's config file using key="SharePointPnPUserAgent" value="PnPRocks"</param>
        /// <param name="spContext">An optional SharePoint client context</param>
        /// <returns>The String value of the result</returns>
        public static String MakeGetRequestForString(String requestUrl,
            String accessToken = null,
            String accept = null,
            Dictionary<string, string> requestHeaders = null,
            Dictionary<string, string> cookies = null,
            int retryCount = 1,
            int delay = 500,
            string userAgent = null,
            ClientContext spContext = null)
        {
            return (MakeHttpRequest<String>("GET",
                requestUrl,
                accessToken: accessToken,
                accept: accept,
                resultPredicate: r => r.Content.ReadAsStringAsync().Result,
                requestHeaders: requestHeaders,
                cookies: cookies,
                retryCount: retryCount,
                delay: delay,
                userAgent: userAgent,
                spContext: spContext
            ));
        }

        /// <summary>
        /// This helper method makes an HTTP GET request and returns the result as a String
        /// </summary>
        /// <param name="requestUrl">The URL of the request</param>
        /// <param name="accessToken">The OAuth 2.0 Access Token for the request, if authorization is required</param>
        /// <param name="accept">The accept header for the response</param>
        /// <param name="referer">The URL Referer for the request</param>
        /// <param name="requestHeaders">A collection of any custom request headers</param>
        /// <param name="cookies">Any request cookies values</param>
        /// <param name="retryCount">Number of times to retry the request</param>
        /// <param name="delay">Milliseconds to wait before retrying the request. The delay will be increased (doubled) every retry</param>
        /// <param name="userAgent">UserAgent string value to insert for this request. You can define this value in your app's config file using key="SharePointPnPUserAgent" value="PnPRocks"</param>
        /// <param name="spContext">An optional SharePoint client context</param>
        /// <returns>The Stream  of the result</returns>
        public static System.IO.Stream MakeGetRequestForStream(string requestUrl,
            string accept,
            string accessToken = null,
            string referer = null,
            Dictionary<string, string> requestHeaders = null,
            Dictionary<string, string> cookies = null,
            int retryCount = 1,
            int delay = 500,
            string userAgent = null,
            ClientContext spContext = null)
        {
            return (MakeHttpRequest<System.IO.Stream>("GET",
                requestUrl,
                accessToken,
                accept: accept,
                referer: referer,
                resultPredicate: r => r.Content.ReadAsStreamAsync().Result,
                requestHeaders: requestHeaders,
                cookies: cookies,
                retryCount: retryCount,
                delay: delay,
                userAgent: userAgent,
                spContext: spContext
            ));
        }

        /// <summary>
        /// This helper method makes an HTTP GET request and returns the result as a String
        /// </summary>
        /// <param name="requestUrl">The URL of the request</param>
        /// <param name="responseHeaders">The response headers of the HTTP request (output argument)</param>
        /// <param name="accessToken">The OAuth 2.0 Access Token for the request, if authorization is required</param>
        /// <param name="accept">The accept header for the response</param>
        /// <param name="requestHeaders">A collection of any custom request headers</param>
        /// <param name="cookies">Any request cookies values</param>
        /// <param name="retryCount">Number of times to retry the request</param>
        /// <param name="delay">Milliseconds to wait before retrying the request. The delay will be increased (doubled) every retry</param>
        /// <param name="userAgent">UserAgent string value to insert for this request. You can define this value in your app's config file using key="SharePointPnPUserAgent" value="PnPRocks"</param>
        /// <param name="spContext">An optional SharePoint client context</param>
        /// <returns>The Stream  of the result</returns>
        public static System.IO.Stream MakeGetRequestForStreamWithResponseHeaders(string requestUrl,
            string accept,
            out HttpResponseHeaders responseHeaders,
            string accessToken = null,
            Dictionary<string, string> requestHeaders = null,
            Dictionary<string, string> cookies = null,
            int retryCount = 1,
            int delay = 500,
            string userAgent = null,
            ClientContext spContext = null)
        {
            return (MakeHttpRequest<System.IO.Stream>("GET",
                requestUrl,
                out responseHeaders,
                accessToken,
                accept: accept,
                resultPredicate: r => r.Content.ReadAsStreamAsync().Result,
                requestHeaders: requestHeaders,
                cookies: cookies,
                retryCount: retryCount,
                delay: delay,
                userAgent: userAgent,
                spContext: spContext
            ));
        }

        /// <summary>
        /// This helper method makes an HTTP POST request without a response
        /// </summary>
        /// <param name="requestUrl">The URL of the request</param>
        /// <param name="accessToken">The OAuth 2.0 Access Token for the request, if authorization is required</param>
        /// <param name="content">The content of the request</param>
        /// <param name="contentType">The content/type of the request</param>
        /// <param name="requestHeaders">A collection of any custom request headers</param>
        /// <param name="cookies">Any request cookies values</param>
        /// <param name="retryCount">Number of times to retry the request</param>
        /// <param name="delay">Milliseconds to wait before retrying the request. The delay will be increased (doubled) every retry</param>
        /// <param name="userAgent">UserAgent string value to insert for this request. You can define this value in your app's config file using key="SharePointPnPUserAgent" value="PnPRocks"</param>
        /// <param name="spContext">An optional SharePoint client context</param>
        public static void MakePostRequest(string requestUrl,
            object content = null,
            string contentType = null,
            string accessToken = null,
            Dictionary<string, string> requestHeaders = null,
            Dictionary<string, string> cookies = null,
            int retryCount = 1,
            int delay = 500,
            string userAgent = null,
            ClientContext spContext = null)
        {
            MakeHttpRequest<string>("POST",
                requestUrl,
                accessToken: accessToken,
                content: content,
                contentType: contentType,
                requestHeaders: requestHeaders,
                cookies: cookies,
                retryCount: retryCount,
                delay: delay,
                userAgent: userAgent,
                spContext: spContext
            );
        }

        /// <summary>
        /// This helper method makes an HTTP POST request and returns the result as a String
        /// </summary>
        /// <param name="requestUrl">The URL of the request</param>
        /// <param name="accessToken">The OAuth 2.0 Access Token for the request, if authorization is required</param>
        /// <param name="content">The content of the request</param>
        /// <param name="contentType">The content/type of the request</param>
        /// <param name="cookies">Any request cookies values</param>
        /// <param name="accept">The value for the accept header in the request</param>
        /// <param name="requestHeaders">A collection of any custom request headers</param>
        /// <param name="retryCount">Number of times to retry the request</param>
        /// <param name="delay">Milliseconds to wait before retrying the request. The delay will be increased (doubled) every retry</param>
        /// <param name="userAgent">UserAgent string value to insert for this request. You can define this value in your app's config file using key="SharePointPnPUserAgent" value="PnPRocks"</param>
        /// <param name="spContext">An optional SharePoint client context</param>
        /// <returns>The String value of the result</returns>
        public static string MakePostRequestForString(string requestUrl,
            object content = null,
            string contentType = null,
            string accessToken = null,
            string accept = null,
            Dictionary<string, string> requestHeaders = null,
            Dictionary<string, string> cookies = null,
            int retryCount = 1,
            int delay = 500,
            string userAgent = null,
            ClientContext spContext = null)
        {
            return (MakeHttpRequest<String>("POST",
                requestUrl,
                accessToken: accessToken,
                content: content,
                contentType: contentType,
                resultPredicate: r => r.Content.ReadAsStringAsync().Result,
                requestHeaders: requestHeaders,
                cookies: cookies,
                accept: accept,
                retryCount: retryCount,
                delay: delay,
                userAgent: userAgent,
                spContext: spContext
            ));
        }

        /// <summary>
        /// This helper method makes an HTTP GET request and returns the result as a String
        /// </summary>
        /// <param name="requestUrl">The URL of the request</param>
        /// <param name="accessToken">The OAuth 2.0 Access Token for the request, if authorization is required</param>
        /// <param name="content">The content of the request</param>
        /// <param name="contentType">The content type for the request</param>
        /// <param name="requestHeaders">A collection of any custom request headers</param>
        /// <param name="cookies">Any request cookies values</param>
        /// <param name="retryCount">Number of times to retry the request</param>
        /// <param name="delay">Milliseconds to wait before retrying the request. The delay will be increased (doubled) every retry</param>
        /// <param name="userAgent">UserAgent string value to insert for this request. You can define this value in your app's config file using key="SharePointPnPUserAgent" value="PnPRocks"</param>
        /// <param name="spContext">An optional SharePoint client context</param>
        /// <returns>The Stream  of the result</returns>
        public static HttpResponseHeaders MakePostRequestForHeaders(string requestUrl,
            object content = null,
            string contentType = null,
            string accessToken = null,
            Dictionary<string, string> requestHeaders = null,
            Dictionary<string, string> cookies = null,
            int retryCount = 1,
            int delay = 500,
            string userAgent = null,
            ClientContext spContext = null)
        {
            return MakeHttpRequest("POST",
                requestUrl,
                accessToken,
                content: content,
                contentType: contentType,
                resultPredicate: response => response.Headers,
                requestHeaders: requestHeaders,
                cookies: cookies,
                retryCount: retryCount,
                delay: delay,
                userAgent: userAgent,
                spContext: spContext
            );
        }

        /// <summary>
        /// This helper method makes an HTTP PUT request without a response
        /// </summary>
        /// <param name="requestUrl">The URL of the request</param>
        /// <param name="accessToken">The OAuth 2.0 Access Token for the request, if authorization is required</param>
        /// <param name="content">The content of the request</param>
        /// <param name="contentType">The content/type of the request</param>
        /// <param name="requestHeaders">A collection of any custom request headers</param>
        /// <param name="cookies">Any request cookies values</param>
        /// <param name="retryCount">Number of times to retry the request</param>
        /// <param name="delay">Milliseconds to wait before retrying the request. The delay will be increased (doubled) every retry</param>
        /// <param name="userAgent">UserAgent string value to insert for this request. You can define this value in your app's config file using key="SharePointPnPUserAgent" value="PnPRocks"</param>
        /// <param name="spContext">An optional SharePoint client context</param>
        public static void MakePutRequest(string requestUrl,
            object content = null,
            string contentType = null,
            string accessToken = null,
            Dictionary<string, string> requestHeaders = null,
            Dictionary<string, string> cookies = null,
            int retryCount = 1,
            int delay = 500,
            string userAgent = null,
            ClientContext spContext = null)
        {
            MakeHttpRequest<string>("PUT",
                requestUrl,
                accessToken: accessToken,
                content: content,
                contentType: contentType,
                requestHeaders: requestHeaders,
                cookies: cookies,
                retryCount: retryCount,
                delay: delay,
                userAgent: userAgent,
                spContext: spContext
            );
        }

        /// <summary>
        /// This helper method makes an HTTP PUT request and returns the result as a String
        /// </summary>
        /// <param name="requestUrl">The URL of the request</param>
        /// <param name="accessToken">The OAuth 2.0 Access Token for the request, if authorization is required</param>
        /// <param name="content">The content of the request</param>
        /// <param name="contentType">The content/type of the request</param>
        /// <param name="requestHeaders">A collection of any custom request headers</param>
        /// <param name="cookies">Any request cookies values</param>
        /// <param name="retryCount">Number of times to retry the request</param>
        /// <param name="delay">Milliseconds to wait before retrying the request. The delay will be increased (doubled) every retry</param>
        /// <param name="userAgent">UserAgent string value to insert for this request. You can define this value in your app's config file using key="SharePointPnPUserAgent" value="PnPRocks"</param>
        /// <param name="spContext">An optional SharePoint client context</param>
        /// <returns>The String value of the result</returns>
        public static string MakePutRequestForString(string requestUrl,
            object content = null,
            string contentType = null,
            string accessToken = null,
            Dictionary<string, string> requestHeaders = null,
            Dictionary<string, string> cookies = null,
            int retryCount = 1,
            int delay = 500,
            string userAgent = null,
            ClientContext spContext = null)
        {
            return (MakeHttpRequest<String>("PUT",
                requestUrl,
                accessToken: accessToken,
                content: content,
                contentType: contentType,
                resultPredicate: r => r.Content.ReadAsStringAsync().Result,
                requestHeaders: requestHeaders,
                cookies: cookies,
                retryCount: retryCount,
                delay: delay,
                userAgent: userAgent,
                spContext: spContext
            ));
        }

        /// <summary>
        /// This helper method makes an HTTP PATCH request and returns the result as a String
        /// </summary>
        /// <param name="requestUrl">The URL of the request</param>
        /// <param name="accessToken">The OAuth 2.0 Access Token for the request, if authorization is required</param>
        /// <param name="content">The content of the request</param>
        /// <param name="contentType">The content/type of the request</param>
        /// <param name="requestHeaders">A collection of any custom request headers</param>
        /// <param name="cookies">Any request cookies values</param>
        /// <param name="retryCount">Number of times to retry the request</param>
        /// <param name="delay">Milliseconds to wait before retrying the request. The delay will be increased (doubled) every retry</param>
        /// <param name="userAgent">UserAgent string value to insert for this request. You can define this value in your app's config file using key="SharePointPnPUserAgent" value="PnPRocks"</param>
        /// <param name="spContext">An optional SharePoint client context</param>
        /// <returns>The String value of the result</returns>
        public static string MakePatchRequestForString(string requestUrl,
            object content = null,
            string contentType = null,
            string accessToken = null,
            Dictionary<string, string> requestHeaders = null,
            Dictionary<string, string> cookies = null,
            int retryCount = 1,
            int delay = 500,
            string userAgent = null,
            ClientContext spContext = null)
        {
            return (MakeHttpRequest<String>("PATCH",
                requestUrl,
                accessToken: accessToken,
                content: content,
                contentType: contentType,
                resultPredicate: r => r.Content.ReadAsStringAsync().Result,
                requestHeaders: requestHeaders,
                cookies: cookies,
                retryCount: retryCount,
                delay: delay,
                userAgent: userAgent,
                spContext: spContext
            ));
        }

        /// <summary>
        /// This helper method makes an HTTP DELETE request
        /// </summary>
        /// <param name="requestUrl">The URL of the request</param>
        /// <param name="accessToken">The OAuth 2.0 Access Token for the request, if authorization is required</param>
        /// <param name="requestHeaders">A collection of any custom request headers</param>
        /// <param name="cookies">Any request cookies values</param>
        /// <param name="retryCount">Number of times to retry the request</param>
        /// <param name="delay">Milliseconds to wait before retrying the request. The delay will be increased (doubled) every retry</param>
        /// <param name="userAgent">UserAgent string value to insert for this request. You can define this value in your app's config file using key="SharePointPnPUserAgent" value="PnPRocks"</param>
        /// <param name="spContext">An optional SharePoint client context</param>
        /// <returns>The String value of the result</returns>
        public static void MakeDeleteRequest(string requestUrl,
            string accessToken = null,
            Dictionary<string, string> requestHeaders = null,
            Dictionary<string, string> cookies = null,
            int retryCount = 1,
            int delay = 500,
            string userAgent = null,
            ClientContext spContext = null)
        {
            MakeHttpRequest<string>("DELETE",
                requestUrl,
                accessToken,
                requestHeaders: requestHeaders,
                cookies: cookies,
                retryCount: retryCount,
                delay: delay,
                userAgent: userAgent,
                spContext: spContext
                );
        }

        /// <summary>
        /// This helper method makes an HTTP request and eventually returns a result
        /// </summary>
        /// <param name="httpMethod">The HTTP method for the request</param>
        /// <param name="requestUrl">The URL of the request</param>
        /// <param name="accessToken">The OAuth 2.0 Access Token for the request, if authorization is required</param>
        /// <param name="accept">The content type of the accepted response</param>
        /// <param name="content">The content of the request</param>
        /// <param name="contentType">The content  type of the request</param>
        /// <param name="referer">The URL Referer for the request</param>
        /// <param name="resultPredicate">The predicate to retrieve the result, if any</param>
        /// <param name="requestHeaders">A collection of any custom request headers</param>
        /// <param name="cookies">Any request cookies values</param>
        /// <param name="retryCount">Number of times to retry the request</param>
        /// <param name="delay">Milliseconds to wait before retrying the request. The delay will be increased (doubled) every retry</param>
        /// <param name="userAgent">UserAgent string value to insert for this request. You can define this value in your app's config file using key="SharePointPnPUserAgent" value="PnPRocks"</param>
        /// <param name="spContext">An optional SharePoint client context</param>
        /// <typeparam name="TResult">The type of the result, if any</typeparam>
        /// <returns>The value of the result, if any</returns>
        private static TResult MakeHttpRequest<TResult>(
            string httpMethod,
            string requestUrl,
            string accessToken = null,
            string accept = null,
            object content = null,
            string contentType = null,
            string referer = null,
            Func<HttpResponseMessage, TResult> resultPredicate = null,
            Dictionary<string, string> requestHeaders = null,
            Dictionary<string, string> cookies = null,
            int retryCount = 1,
            int delay = 500,
            string userAgent = null,
            ClientContext spContext = null)
        {
            HttpResponseHeaders responseHeaders;
            return (MakeHttpRequest<TResult>(httpMethod,
                requestUrl,
                out responseHeaders,
                accessToken: accessToken,
                accept: accept,
                content: content,
                contentType: contentType,
                referer: referer,
                resultPredicate: resultPredicate,
                requestHeaders: requestHeaders,
                cookies: cookies,
                retryCount: retryCount,
                delay: delay,
                userAgent: userAgent,
                spContext: spContext
                ));
        }

        /// <summary>
        /// This helper method makes an HTTP request and eventually returns a result
        /// </summary>
        /// <param name="httpMethod">The HTTP method for the request</param>
        /// <param name="requestUrl">The URL of the request</param>
        /// <param name="responseHeaders">The response headers of the HTTP request (output argument)</param>
        /// <param name="accessToken">The OAuth 2.0 Access Token for the request, if authorization is required</param>
        /// <param name="accept">The content type of the accepted response</param>
        /// <param name="content">The content of the request</param>
        /// <param name="contentType">The content  type of the request</param>
        /// <param name="referer">The URL Referer for the request</param>
        /// <param name="resultPredicate">The predicate to retrieve the result, if any</param>
        /// <param name="requestHeaders">A collection of any custom request headers</param>
        /// <param name="cookies">Any request cookies values</param>
        /// <param name="retryCount">Number of times to retry the request</param>
        /// <param name="delay">Milliseconds to wait before retrying the request. The delay will be increased (doubled) every retry</param>
        /// <param name="userAgent">UserAgent string value to insert for this request. You can define this value in your app's config file using key="SharePointPnPUserAgent" value="PnPRocks"</param>
        /// <param name="spContext">An optional SharePoint client context</param>
        /// <typeparam name="TResult">The type of the result, if any</typeparam>
        /// <returns>The value of the result, if any</returns>
        internal static TResult MakeHttpRequest<TResult>(
            string httpMethod,
            string requestUrl,
            out HttpResponseHeaders responseHeaders,
            string accessToken = null,
            string accept = null,
            object content = null,
            string contentType = null,
            string referer = null,
            Func<HttpResponseMessage, TResult> resultPredicate = null,
            Dictionary<string, string> requestHeaders = null,
            Dictionary<string, string> cookies = null,
            int retryCount = 1,
            int delay = 500,
            string userAgent = null,
            ClientContext spContext = null)
        {
            var result = MakeHttpRequestAsync(httpMethod, requestUrl, accessToken, accept, content, contentType, referer, resultPredicate, requestHeaders, cookies, retryCount, delay, userAgent, spContext).Result;
            responseHeaders = result.ResponseHeaders;
            return (result.Result);
        }

        /// <summary>
        /// This helper method makes an HTTP request and eventually returns a result
        /// </summary>
        /// <param name="httpMethod">The HTTP method for the request</param>
        /// <param name="requestUrl">The URL of the request</param>
        /// <param name="responseHeaders">The response headers of the HTTP request (output argument)</param>
        /// <param name="accessToken">The OAuth 2.0 Access Token for the request, if authorization is required</param>
        /// <param name="accept">The content type of the accepted response</param>
        /// <param name="content">The content of the request</param>
        /// <param name="contentType">The content  type of the request</param>
        /// <param name="referer">The URL Referer for the request</param>
        /// <param name="resultPredicate">The predicate to retrieve the result, if any</param>
        /// <param name="requestHeaders">A collection of any custom request headers</param>
        /// <param name="cookies">Any request cookies values</param>
        /// <param name="retryCount">Number of times to retry the request</param>
        /// <param name="delay">Milliseconds to wait before retrying the request. The delay will be increased (doubled) every retry</param>
        /// <param name="userAgent">UserAgent string value to insert for this request. You can define this value in your app's config file using key="SharePointPnPUserAgent" value="PnPRocks"</param>
        /// <param name="spContext">An optional SharePoint client context</param>
        /// <typeparam name="TResult">The type of the result, if any</typeparam>
        /// <returns>The value of the result, if any</returns>
        internal static async Task<HttpResult<TResult>> MakeHttpRequestAsync<TResult>(
            string httpMethod,
            string requestUrl,
            string accessToken = null,
            string accept = null,
            object content = null,
            string contentType = null,
            string referer = null,
            Func<HttpResponseMessage, TResult> resultPredicate = null,
            Dictionary<string, string> requestHeaders = null,
            Dictionary<string, string> cookies = null,
            int retryCount = 1,
            int delay = 500,
            string userAgent = null,
            ClientContext spContext = null)
        {
            //HttpClient client = HttpHelper.httpClient;
            HttpClient client;

            // Define whether to use the default HttpClient object
            if (spContext != null)
            {
#pragma warning disable CA2000 // Dispose objects before losing scope
                client = PnPHttpClient.Instance.GetHttpClient(spContext);
#pragma warning restore CA2000 // Dispose objects before losing scope
            }
            else
            {
                client = PnPHttpClient.Instance.GetHttpClient();
            }

            // Prepare the variable to hold the result, if any
            TResult result = default;
            HttpResponseHeaders responseHeaders = null;

            if (!string.IsNullOrEmpty(referer))
            {
                client.DefaultRequestHeaders.Referrer = new Uri(referer);
            }

            // If there is an accept argument, set the corresponding HTTP header
            if (!string.IsNullOrEmpty(accept))
            {
                client.DefaultRequestHeaders.Accept.Clear();
                client.DefaultRequestHeaders.Accept.Add(
                    new MediaTypeWithQualityHeaderValue(accept));
            }

            // Process any additional custom request headers
            if (requestHeaders != null)
            {
                foreach (var requestHeader in requestHeaders)
                {
                    client.DefaultRequestHeaders.Add(requestHeader.Key, requestHeader.Value);
                }
            }

            // Prepare the content of the request, if any
            HttpContent requestContent = null;
            System.IO.Stream streamContent = content as System.IO.Stream;
            if (streamContent != null)
            {
                requestContent = new StreamContent(streamContent);
                requestContent.Headers.ContentType = new MediaTypeHeaderValue(contentType);
            }
            else if (content != null)
            {
                var jsonString = content is string
                    ? content.ToString()
                    : JsonConvert.SerializeObject(content, Formatting.None, new JsonSerializerSettings
                    {
                        NullValueHandling = NullValueHandling.Ignore,
                        ContractResolver = new ODataBindJsonResolver(),

                    });
                requestContent = new StringContent(jsonString, Encoding.UTF8, contentType);
            }

            // Prepare the HTTP request message with the proper HTTP method
            using (HttpRequestMessage request = new HttpRequestMessage(new HttpMethod(httpMethod), requestUrl))
            {
                // Set the request content, if any
                if (requestContent != null)
                {
                    request.Content = requestContent;
                }

                if (spContext != null)
                {
                    PnPHttpClient.AuthenticateRequestAsync(request, spContext).GetAwaiter().GetResult();
                }
                else
                {
                    PnPHttpClient.AuthenticateRequest(request, accessToken);
                }

                // Fire the HTTP request
                HttpResponseMessage response = await client.SendAsync(request);

                if (response.IsSuccessStatusCode)
                {
                    // If the response is Success and there is a
                    // predicate to retrieve the result, invoke it
                    if (resultPredicate != null)
                    {
                        result = resultPredicate(response);
                    }

                    // Get any response header and put it in the answer
                    responseHeaders = response.Headers;
                }
                else
                {
                    throw new ApplicationException(
                        string.Format("Exception while invoking endpoint {0}.", requestUrl),
                        new Exception(response.Content.ReadAsStringAsync().Result));
                }
            }

            return new HttpResult<TResult>(result, responseHeaders);
        }

    }
}
