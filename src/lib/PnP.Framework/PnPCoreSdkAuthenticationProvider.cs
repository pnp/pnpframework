using Microsoft.SharePoint.Client;
using PnP.Core.Services;
using PnP.Framework.Modernization.Utilities;
using System;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Threading.Tasks;

namespace PnP.Framework
{
    internal class PnPCoreSdkAuthenticationProvider : ILegacyAuthenticationProvider
    {
        private readonly ClientContext clientContext;

        internal PnPCoreSdkAuthenticationProvider(ClientContext context)
        {
            clientContext = context ?? throw new ArgumentNullException(nameof(context));
        }

        public async Task AuthenticateRequestAsync(Uri resource, HttpRequestMessage request)
        {
            if (resource == null)
            {
                throw new ArgumentNullException(nameof(resource));
            }

            if (request == null)
            {
                throw new ArgumentNullException(nameof(request));
            }

            request.Headers.Authorization = new AuthenticationHeaderValue("bearer",
                await GetAccessTokenAsync(resource).ConfigureAwait(false));
        }

#pragma warning disable CS1998 // Async method lacks 'await' operators and will run synchronously
        public async Task<string> GetAccessTokenAsync(Uri resource, string[] scopes)
#pragma warning restore CS1998 // Async method lacks 'await' operators and will run synchronously
        {
            return GetAccessToken(resource);
        }

#pragma warning disable CS1998 // Async method lacks 'await' operators and will run synchronously
        public async Task<string> GetAccessTokenAsync(Uri resource)
#pragma warning restore CS1998 // Async method lacks 'await' operators and will run synchronously
        {
            return GetAccessToken(resource);
        }

        private string GetAccessToken(Uri resource)
        {
            string accessToken = null;
            if (resource == new Uri(clientContext.Url))
            {
                accessToken = clientContext.GetAccessToken();
            }
            else
            {
                using (var context = clientContext.Clone(resource))
                {
                    accessToken = context.GetAccessToken();
                }
            }

            return accessToken;
        }

        public string GetCookieHeader(Uri targetUrl)
        {
            var cookieManager = new CookieManager();
            var cookieContainer = cookieManager.GetCookies(clientContext);
            if (cookieContainer == null)
            {
                throw new InvalidOperationException("Unable to access CookieManager for current ClientContext instance");
            }    

            return cookieContainer.GetCookieHeader(targetUrl);
        }
    }
}
