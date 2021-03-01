using Microsoft.Extensions.DependencyInjection;
using Microsoft.SharePoint.Client;
using PnP.Framework.Utilities;
using System;
using System.Collections.Concurrent;
using System.Configuration;
using System.Net;
using System.Net.Http;
using System.Threading.Tasks;

namespace PnP.Framework.Http
{
    /// <summary>
    /// Static class holding HttpClient references, needs to be static to avoid port exhaustion scenarios
    /// </summary>
    public class PnPHttpClient
    {
        private const string PnPHttpClientName = "PnPHttpClient";
        private static readonly Lazy<PnPHttpClient> _lazyInstance = new Lazy<PnPHttpClient>(() => new PnPHttpClient(), true);
        private ServiceProvider serviceProvider;
        private static readonly ConcurrentDictionary<string, HttpClientHandler> credentialsHttpClients = new ConcurrentDictionary<string, HttpClientHandler>();

        private PnPHttpClient()
        {
            BuildServiceFactory();
        }

        public static PnPHttpClient Instance
        {
            get
            {
                return _lazyInstance.Value;
            }
        }

        public HttpClient GetHttpClient(ClientContext context)
        {
            var factory = serviceProvider.GetRequiredService<IHttpClientFactory>();

            if (context.Credentials is NetworkCredential networkCredential)
            {
                string cacheKey = networkCredential.UserName;

                if (string.IsNullOrEmpty(cacheKey))
                {
                    cacheKey = CredentialCache.DefaultNetworkCredentials.UserName;
                }

                // The HttpClientHandler is the one managing the network connections and holds the resources and as
                // such we're caching this one for on-prem usage scenarions (for page transformation)
                if (credentialsHttpClients.TryGetValue(cacheKey, out HttpClientHandler cachedHttpHandler))
                {
                    // No need to dispose HttpClient, the IDisposable is purely there to trigger the 
                    // dispose of the created HttpClientHandler
                    return new HttpClient(cachedHttpHandler);
                }
                else
                {
                    // Create a new handler, do not dispose it since we're caching it
                    var handler = new HttpClientHandler
                    {
                        Credentials = context.Credentials
                    };

                    credentialsHttpClients.TryAdd(cacheKey, handler);

                    // No need to dispose HttpClient, the IDisposable is purely there to trigger the 
                    // dispose of the created HttpClientHandler
                    return new HttpClient(handler);
                }
            }
            else
            {
                // Let the HttpClientFactory handle things
                return factory.CreateClient(PnPHttpClientName);
            }
        }

        public HttpClient GetHttpClient()
        {
            var factory = serviceProvider.GetRequiredService<IHttpClientFactory>();
            return factory.CreateClient(PnPHttpClientName);
        }

        public static async Task AuthenticateRequestAsync(HttpRequestMessage request, ClientContext context)
        {
            var accessToken = context.GetAccessToken();

            if (!string.IsNullOrEmpty(accessToken))
            {
                request.Headers.Authorization = new System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", accessToken);
            }
            else 
            {
                var cookieContainer = context.GetAuthenticationCookies();
                if (cookieContainer != null)
                {
                    request.Headers.Add("Cookie", cookieContainer.GetCookieHeader(new Uri(context.Url)));
                    if (request.Method != HttpMethod.Get)
                    {
                        request.Headers.Add("X-RequestDigest", await context.GetRequestDigestAsync(cookieContainer).ConfigureAwait(false));
                    }                    
                }
                else if (context.Credentials is NetworkCredential)
                {
                    // No need to add credentials as these are already provided via the selected HttpClient/HttpClientHandler
                    if (request.Method != HttpMethod.Get)
                    {
                        request.Headers.Add("X-RequestDigest", await context.GetOnPremisesRequestDigestAsync().ConfigureAwait(false));
                    }
                }
            }
        }

        public static void AuthenticateRequest(HttpRequestMessage request, string accessToken)
        {
            if (!string.IsNullOrEmpty(accessToken))
            {
                request.Headers.Authorization = new System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", accessToken);
            }
        }

        private void BuildServiceFactory()
        {
            // Use TLS 1.2 as default connection
            ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12 | SecurityProtocolType.Tls11 | SecurityProtocolType.Tls;

            // Create container
            var serviceCollection = new ServiceCollection();

            // Add http handlers
            AddHttpHandlers(serviceCollection);

            // get User Agent String
            string userAgentFromConfig = null;
            try
            {
                userAgentFromConfig = ConfigurationManager.AppSettings["SharePointPnPUserAgent"];
            }
            catch // throws exception if being called from a .NET Standard 2.0 application
            {

            }
            if (string.IsNullOrWhiteSpace(userAgentFromConfig))
            {
                userAgentFromConfig = Environment.GetEnvironmentVariable("SharePointPnPUserAgent", EnvironmentVariableTarget.Process);
            }

            // Add http clients
            AddHttpClients(serviceCollection, userAgentFromConfig);

            // Build the container
            serviceProvider = serviceCollection.BuildServiceProvider();
        }

        private static IServiceCollection AddHttpClients(IServiceCollection collection, string UserAgent = null)
        {
            collection.AddHttpClient(PnPHttpClientName, config =>
            {
                if (string.IsNullOrWhiteSpace(UserAgent))
                {
                    config.DefaultRequestHeaders.UserAgent.TryParseAdd(PnPCoreUtilities.PnPCoreUserAgent);
                }
                else
                {
                    config.DefaultRequestHeaders.UserAgent.TryParseAdd(UserAgent);
                }
            })
            .AddHttpMessageHandler<RetryHandler>()
            // We use cookies by adding them to the header which works great when used from Core framework,
            // however when running the .NET Standard 2.0 version from .NET Framework we explicetely have to
            // tell the http client to not use the default (empty) cookie container
            .ConfigurePrimaryHttpMessageHandler(() => new HttpClientHandler()
            {
                UseCookies = false
            });

            return collection;
        }

        private static IServiceCollection AddHttpHandlers(IServiceCollection collection)
        {
            // Use transient for the DelegatingHandlers
            // https://stackoverflow.com/questions/53223411/httpclient-delegatinghandler-unexpected-life-cycle
            collection.AddTransient<RetryHandler, RetryHandler>();

            return collection;
        }
    }
}
