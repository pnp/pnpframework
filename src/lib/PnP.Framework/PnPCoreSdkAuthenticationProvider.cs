using Microsoft.SharePoint.Client;
using PnP.Core.Services;
using PnP.Framework.Modernization.Utilities;
using System;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Threading.Tasks;

namespace PnP.Framework
{
    internal class PnPCoreSdkAuthenticationProvider : ILegacyAuthenticationProvider
    {
        private readonly ClientContext clientContext;
        private readonly string userAgent;
        private CookieContainer cookieContainer;

        internal PnPCoreSdkAuthenticationProvider(ClientContext context)
        {
            clientContext = context ?? throw new ArgumentNullException(nameof(context));
        }

        internal PnPCoreSdkAuthenticationProvider(ClientContext context, string userAgent)
        {
            clientContext = context ?? throw new ArgumentNullException(nameof(context));
            this.userAgent = userAgent ?? throw new ArgumentNullException(nameof(userAgent));
        }

        public PnPCoreSdkAuthenticationProvider()
        {
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
            if (!string.IsNullOrEmpty(userAgent))
            {
                request.Headers.Remove("User-Agent");
                request.Headers.Add("User-Agent", userAgent);
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
            if (cookieContainer == null)
            {
                throw new InvalidOperationException("Unable to access CookieContainer for current ClientContext instance");
            }

            return cookieContainer.GetCookieHeader(targetUrl);
        }

        public string GetRequestDigest()
        {
            if (cookieContainer == null)
            {
                throw new InvalidOperationException("Unable to access CookieContainer for current ClientContext instance");
            }

            return clientContext.GetRequestDigestAsync(cookieContainer).GetAwaiter().GetResult();
        }

        public bool RequiresCookieAuthentication
        {
            get
            {
                var contextSettings = clientContext.GetContextSettings();
                if (contextSettings != null && contextSettings.Type == Utilities.Context.ClientContextType.Cookie)
                {
                    if (cookieContainer == null)
                    {
                        var cookieManager = new CookieManager();
                        cookieContainer = cookieManager.GetCookies(clientContext);
                    }
                    return true;
                }
                else
                {
                    return false;
                }
            }
        }
    }
}
