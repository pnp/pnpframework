using Microsoft.Graph;
using PnP.Framework.Utilities;
using System;
using System.Net.Http;
using System.Threading;
using System.Threading.Tasks;

namespace PnP.Framework.Graph
{
    ///<summary>
    /// Class that deals with PnPHttpProvider methods
    ///</summary>  
    public class PnPHttpProvider : HttpProvider, IHttpProvider
    {
        private readonly string _userAgent;
        private readonly PnPHttpRetryHandler _retryHandler;

        /// <summary>
        /// Constructor for the PnPHttpProvider class
        /// </summary>
        /// <param name="retryCount">Maximum retry Count</param>
        /// <param name="delay">Delay Time</param>
        /// <param name="userAgent">User-Agent string to set</param>
        public PnPHttpProvider(int retryCount = 10, int delay = 500, string userAgent = null) : base()
        {
            if (retryCount <= 0)
                throw new ArgumentException("Provide a retry count greater than zero.");

            if (delay <= 0)
                throw new ArgumentException("Provide a delay greater than zero.");

            this._userAgent = userAgent;
            this._retryHandler = new PnPHttpRetryHandler(retryCount, delay);
        }

        /// <summary>
        /// Custom implementation of the IHttpProvider.SendAsync method to handle retry logic
        /// </summary>
        /// <param name="request">The HTTP Request Message</param>
        /// <param name="completionOption">The completion option</param>
        /// <param name="cancellationToken">The cancellation token</param>
        /// <returns>The result of the asynchronous request</returns>
        /// <remarks>See here for further details: https://graph.microsoft.io/en-us/docs/overview/errors</remarks>
        async Task<HttpResponseMessage> IHttpProvider.SendAsync(HttpRequestMessage request, HttpCompletionOption completionOption, CancellationToken cancellationToken)
        {
            // Add the PnP User Agent string
            request.Headers.UserAgent.TryParseAdd(string.IsNullOrEmpty(_userAgent) ? $"{PnPCoreUtilities.PnPCoreUserAgent}" : _userAgent);

            return await _retryHandler.SendRetryAsync(request, (r) => base.SendAsync(r, completionOption, cancellationToken), cancellationToken);
        }
    }
}
