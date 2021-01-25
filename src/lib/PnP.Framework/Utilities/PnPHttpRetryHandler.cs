using PnP.Framework.Diagnostics;
using System;
using System.Collections.Generic;
using System.IO;
using System.Net;
using System.Net.Http;
using System.Threading;
using System.Threading.Tasks;

namespace PnP.Framework.Utilities
{
    /// <summary>
    /// Internal helper class which implements retry mechanismn on throttling
    /// </summary>
    internal class PnPHttpRetryHandler
    {
        readonly int retryCount;
        readonly int delay;

        /// <summary>
        /// Constructor with PnPHttpRetryHandler
        /// </summary>
        /// <param name="retryCount">Number of retries, defaults to 10</param>
        /// <param name="delay">Incremental delay increase in milliseconds</param>
        public PnPHttpRetryHandler(int retryCount = 10, int delay = 500)
        {
            this.retryCount = retryCount;
            this.delay = delay;
        }

        public delegate Task<HttpResponseMessage> HttpPerformRequest(HttpRequestMessage request);

        /// <summary>
        /// Perform async http request, retry if server is unavailable or 
        /// </summary>
        /// <param name="request">Http request to execute</param>
        /// <param name="performRequest">Delegate that performs the request</param>
        /// <param name="cancellationToken">cancellation token</param>
        /// <returns>Response object from http request</returns>
        public async Task<HttpResponseMessage> SendRetryAsync(HttpRequestMessage request, HttpPerformRequest performRequest, CancellationToken cancellationToken)
        {
            // Retry logic variables
            int backoffInterval = this.delay;

            // Loop until we need to retry
            for (int retryAttempts = 0; retryAttempts < this.retryCount; retryAttempts++)
            {
                try
                {
                    // Make the request
                    var httpResponse = await performRequest(request);

                    // return if we got no response or there is no need to retry
                    if (httpResponse == null || httpResponse.StatusCode != (HttpStatusCode)429 && httpResponse.StatusCode != (HttpStatusCode)503)
                    {
                        return httpResponse;
                    }
                }
                // Or handle any ServiceException
                catch (Exception ex)
                {
                    // Check if the is an InnerException
                    // And if it is a WebException
                    if (ex.InnerException is WebException wex && wex.Response is HttpWebResponse response)
                    {
                        // Check if request was throttled - http status code 429
                        // Check is request failed due to server unavailable - http status code 503
                        if (response.StatusCode != (HttpStatusCode)429 && response.StatusCode != (HttpStatusCode)503)
                        {
                            Log.Error(Constants.LOGGING_SOURCE, CoreResources.GraphExtensions_SendAsyncRetryException, wex.ToString());
                            throw;
                        }

                        // else: retry!
                    }
                    else
                    {
                        throw;
                    }
                }

                Log.Warning(Constants.LOGGING_SOURCE, CoreResources.GraphExtensions_SendAsyncRetry, backoffInterval);

                // Add delay for retry
                Task.Delay(backoffInterval, cancellationToken).Wait(cancellationToken);

                // increase delay.
                backoffInterval = backoffInterval * 2;

                // clone request
                var oldRequest = request;
                request = oldRequest.CloneRequest();
                oldRequest.Dispose();
            }

            throw new Microsoft.SharePoint.Client.ClientContextExtensions.MaximumRetryAttemptedException($"Maximum retry attempts {this.retryCount}, has be attempted.");
        }
    }

    //Reference: https://stackoverflow.com/questions/18000583/re-send-httprequestmessage-exception/18014515#18014515
    internal static class PnPHttpRequestCloneExtension
    {
        public static HttpRequestMessage CloneRequest(this HttpRequestMessage request)
        {
            var clone = new HttpRequestMessage(request.Method, request.RequestUri)
            {
                Content = request.Content.CloneRequest(),
                Version = request.Version
            };
#pragma warning disable CS0618
            foreach (KeyValuePair<string, object> prop in request.Properties)
            {
                clone.Properties.Add(prop);
            }
#pragma warning restore CS0618
            foreach (KeyValuePair<string, IEnumerable<string>> header in request.Headers)
            {
                clone.Headers.TryAddWithoutValidation(header.Key, header.Value);
            }

            return clone;
        }

        public static HttpContent CloneRequest(this HttpContent content)
        {
            if (content == null) return null;

            var ms = new MemoryStream();
            content.CopyToAsync(ms).ConfigureAwait(false).GetAwaiter().GetResult();
            ms.Position = 0;

            var clone = new StreamContent(ms);
            foreach (KeyValuePair<string, IEnumerable<string>> header in content.Headers)
            {
                clone.Headers.Add(header.Key, header.Value);
            }
            return clone;
        }
    }
}
