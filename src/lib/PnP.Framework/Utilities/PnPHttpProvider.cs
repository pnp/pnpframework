using System;
using System.Net;
using System.Net.Http;
using System.Threading;
using System.Threading.Tasks;

namespace PnP.Framework.Utilities
{
    /// <summary>	
    /// PnP http client which implements setting of User-Agent + retry mechanismn on throttling	
    /// </summary>	
    [Obsolete("Please use PnPHttpClient.Instance.GetHttpClient and PnPHttpClient.AuthenticateRequestAsync instead.")]
    public class PnPHttpProvider : HttpClient
    {
        private readonly string userAgent;
        private readonly PnPHttpRetryHandler retryHandler;

        /// <summary>	
        /// Constructor without HttpMessageHandler	
        /// </summary>	
        /// <param name="retryCount">Number of retries, defaults to 10</param>	
        /// <param name="delay">Incremental delay increase in milliseconds</param>	
        /// <param name="userAgent">User-Agent string to set</param>	
        public PnPHttpProvider(int retryCount = 10, int delay = 500, string userAgent = null) : this(new HttpClientHandler(), retryCount, delay, userAgent)
        {
        }

        /// <summary>	
        /// Constructor with HttpMessageHandler	
        /// </summary>	
        /// <param name="innerHandler">HttpMessageHandler instance to pass along</param>	
        /// <param name="retryCount">Number of retries, defaults to 10</param>	
        /// <param name="delay">Incremental delay increase in milliseconds</param>	
        /// <param name="userAgent">User-Agent string to set</param>	
        public PnPHttpProvider(HttpMessageHandler innerHandler, int retryCount = 10, int delay = 500, string userAgent = null) : this(innerHandler, false, retryCount, delay, userAgent)
        {
        }

        /// <summary>	
        /// Constructor with HttpMessageHandler	
        /// </summary>	
        /// <param name="innerHandler">HttpMessageHandler instance to pass along</param>	
        /// <param name="retryCount">Number of retries, defaults to 10</param>	
        /// <param name="delay">Incremental delay increase in milliseconds</param>	
        /// <param name="userAgent">User-Agent string to set</param>	
        /// <param name="disposeHandler">Declares whether to automatically dispose the internal HttpHandler instance</param>	
        public PnPHttpProvider(HttpMessageHandler innerHandler, bool disposeHandler, int retryCount = 10, int delay = 500, string userAgent = null) : base(innerHandler, disposeHandler)
        {
            this.userAgent = userAgent;

#if !NET9_0
            // Use TLS 1.2 as default connection	
            ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12 | SecurityProtocolType.Tls11 | SecurityProtocolType.Tls;
#endif

            this.retryHandler = new PnPHttpRetryHandler(retryCount, delay);
        }

        /// <summary>	
        /// Perform async http request	
        /// </summary>	
        /// <param name="request">Http request to execute</param>	
        /// <param name="cancellationToken">cancellation token</param>	
        /// <returns>Response object from http request</returns>	
        public async override Task<HttpResponseMessage> SendAsync(HttpRequestMessage request, CancellationToken cancellationToken)
        {
            // Add the PnP User Agent string	
            request.Headers.UserAgent.TryParseAdd(string.IsNullOrEmpty(userAgent) ? $"{PnPCoreUtilities.PnPCoreUserAgent}" : userAgent);

            return await retryHandler.SendRetryAsync(request, (r) => base.SendAsync(r, cancellationToken), cancellationToken);
        }
    }
}