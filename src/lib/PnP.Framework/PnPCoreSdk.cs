using Microsoft.Extensions.DependencyInjection;
using Microsoft.SharePoint.Client;
using PnP.Core.Services;
using PnP.Framework.Utilities.Context;
using PnP.Framework.Utilities.PnPSdk;
using System;
using System.Threading;
using System.Threading.Tasks;

namespace PnP.Framework
{
    /// <summary>
    /// Class that implements interop between PnP Framework and PnP Core SDK
    /// </summary>
    public class PnPCoreSdk
    {

        private static readonly Lazy<PnPCoreSdk> _lazyInstance = new Lazy<PnPCoreSdk>(() => new PnPCoreSdk(), true);
        internal IPnPContextFactory pnpContextFactoryCache;
        private static readonly SemaphoreSlim semaphoreSlimFactory = new SemaphoreSlim(1);
        internal static ILegacyAuthenticationProviderFactory AuthenticationProviderFactory { get; set; } = new PnPCoreSdkAuthenticationProviderFactory();
        internal static event EventHandler<IServiceCollection> OnDIContainerBuilding;

        /// <summary>
        /// Provides the singleton instance of th entity manager
        /// </summary>
        public static PnPCoreSdk Instance
        {
            get
            {
                return _lazyInstance.Value;
            }
        }

        /// <summary>
        /// Private constructor since this is a singleton
        /// </summary>
        private PnPCoreSdk()
        {
        }

        /// <summary>
        /// Get's a PnPContext from a CSOM ClientContext
        /// </summary>
        /// <param name="context">CSOM ClientContext</param>
        /// <param name="existingFactory">An existing factory to use for PnPContext creation, instead of an internal one.</param>
        /// <returns>The equivalent PnPContext</returns>
        public async Task<PnPContext> GetPnPContextAsync(ClientContext context, IPnPContextFactory existingFactory = null)
        {
            Uri ctxUri = new Uri(context.Url);

            var ctxSettings = context.GetContextSettings();

            if (ctxSettings != null && ctxSettings.Type == Utilities.Context.ClientContextType.PnPCoreSdk && ctxSettings.AuthenticationManager != null)
            {
                var pnpContext = ctxSettings.AuthenticationManager.PnPCoreContext;
                if (pnpContext != null && pnpContext.Uri == ctxUri)
                {
                    return pnpContext;
                }
                else
                {
                    var iAuthProvider = ctxSettings.AuthenticationManager.PnPCoreAuthenticationProvider;
                    if (iAuthProvider != null)
                    {
                        IPnPContextFactory factory0;
                        if (existingFactory != null)
                        {
                            // use the provided factory for all upcoming PnPContext creations, also the ones driven internally from PnP Framework
                            pnpContextFactoryCache = existingFactory;
                            factory0 = existingFactory;
                        }
                        else
                        {
                            factory0 = BuildContextFactory(ctxSettings);                            
                        }
                        
                        return await factory0.CreateAsync(ctxUri, iAuthProvider).ConfigureAwait(false);

                    }
                }
            }

            if (existingFactory != null)
            {
                // use the provided factory for all upcoming PnPContext creations, also the ones driven internally from PnP Framework
                pnpContextFactoryCache = existingFactory;                
            }
            
            var factory = existingFactory ?? BuildContextFactory(ctxSettings);
            return await factory.CreateAsync(ctxUri, AuthenticationProviderFactory.GetAuthenticationProvider(context)).ConfigureAwait(false);
        }

        /// <summary>
        /// Get's a PnPContext from a CSOM ClientContext
        /// </summary>
        /// <param name="context">CSOM ClientContext</param>
        /// <param name="existingFactory">An existing factory to use for PnPContext creation, instead of an internal one.</param>
        /// <returns>The equivalent PnPContext</returns>
        public PnPContext GetPnPContext(ClientContext context, IPnPContextFactory existingFactory = null)
        {
            return GetPnPContextAsync(context, existingFactory).GetAwaiter().GetResult();
        }

        private IPnPContextFactory BuildContextFactory(ClientContextSettings clientContextSettings)
        {
            try
            {
                // Ensure there's only one context factory building happening at any given time
                semaphoreSlimFactory.Wait();

                // Return the factory from cache if we already have one
                if (pnpContextFactoryCache != null)
                {
                    return pnpContextFactoryCache;
                }

                // Build the service collection and load PnP Core SDK
                IServiceCollection services = new ServiceCollection();

                string environmentToUse = "Production";
                string azureADLoginAuthority = "login.microsoftonline.com";
                string microsoftGraphAuthority = "graph.microsoft.com";
                if (clientContextSettings != null)
                {
                    environmentToUse = MapAzureEnvironmentToMicrosoft365Environment(clientContextSettings.Environment.ToString());

                    if (environmentToUse == "Custom" && clientContextSettings.AuthenticationManager != null)
                    {
                        string azureAdLoginEndPointForCustomConfiguration = clientContextSettings.AuthenticationManager.GetAzureAdLoginEndPointForCustomAzureEnvironmentConfiguration();
                        if (!string.IsNullOrEmpty(azureAdLoginEndPointForCustomConfiguration))
                        { 
                            if (Uri.TryCreate(azureAdLoginEndPointForCustomConfiguration, UriKind.Absolute, out Uri azureAdLoginEndPointUri))
                            {
                                azureADLoginAuthority = azureAdLoginEndPointUri.Authority;
                            }
                        }

                        string graphLoginEndPointForCustomConfiguration = clientContextSettings.AuthenticationManager.GetGraphEndPointForCustomAzureEnvironmentConfiguration();
                        if (!string.IsNullOrEmpty(graphLoginEndPointForCustomConfiguration))
                        {
                            microsoftGraphAuthority = graphLoginEndPointForCustomConfiguration;
                        }
                    }
                }

                if (environmentToUse == "Custom")
                {
                    // To increase coverage of solutions providing tokens without graph scopes we turn of graphfirst for PnPContext created from PnP Framework                
                    services = services.AddPnPCore(options =>
                    {
                        options.PnPContext.GraphFirst = false;
                        options.Environment = environmentToUse;
                        options.AzureADLoginAuthority = azureADLoginAuthority;
                        options.MicrosoftGraphAuthority = microsoftGraphAuthority;
                    }).Services;
                }
                else
                {
                    // To increase coverage of solutions providing tokens without graph scopes we turn of graphfirst for PnPContext created from PnP Framework                
                    services = services.AddPnPCore(options =>
                    {
                        options.PnPContext.GraphFirst = false;
                        options.Environment = environmentToUse;
                    }).Services;
                }

                // Enables to plug in additional services into this service container
                if (OnDIContainerBuilding != null)
                {
                    OnDIContainerBuilding.Invoke(this, services);
                }

                var serviceProvider = services.BuildServiceProvider();

                // Get a PnP context factory
                var pnpContextFactory = serviceProvider.GetRequiredService<IPnPContextFactory>();

                // Chache the factory before returning it
                if (pnpContextFactoryCache == null)
                {
                    pnpContextFactoryCache = pnpContextFactory;
                }

                return pnpContextFactory;
            }
            finally
            {
                semaphoreSlimFactory.Release();
            }

        }

        /// <summary>
        /// Returns a CSOM ClientContext for a given PnP Core SDK context
        /// </summary>
        /// <param name="pnpContext">The PnP Core SDK context</param>
        /// <returns>The equivalent CSOM ClientContext</returns>
        public async Task<ClientContext> GetClientContextAsync(PnPContext pnpContext)
        {
#pragma warning disable CA2000 // Dispose objects before losing scope
            AuthenticationManager authManager = AuthenticationManager.CreateWithPnPCoreSdk(pnpContext);
#pragma warning restore CA2000 // Dispose objects before losing scope

            var ctx = await authManager.GetContextAsync(pnpContext.Uri.ToString()).ConfigureAwait(false);
            var ctxSettings = ctx.GetContextSettings();
            ctxSettings.Type = Utilities.Context.ClientContextType.PnPCoreSdk;
            ctxSettings.AuthenticationManager = authManager; //otherwise GetAccessToken would not work for example
            ctx.AddContextSettings(ctxSettings);
            return ctx;
        }

        /// <summary>
        /// Returns a CSOM ClientContext for a given PnP Core SDK context
        /// </summary>
        /// <param name="pnpContext">The PnP Core SDK context</param>
        /// <returns>The equivalent CSOM ClientContext</returns>
        public ClientContext GetClientContext(PnPContext pnpContext)
        {
            return GetClientContextAsync(pnpContext).GetAwaiter().GetResult();
        }

        private static string MapAzureEnvironmentToMicrosoft365Environment(string azureEnvironment) 
        { 
            if (azureEnvironment == "PPE")
            {
                return "PreProduction";
            }
            else
            {
                return azureEnvironment;
            }
        }

    }
}
