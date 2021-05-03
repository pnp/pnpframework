using Microsoft.Extensions.DependencyInjection;
using Microsoft.SharePoint.Client;
using PnP.Core.Services;
using PnP.Framework.Utilities.PnPSdk;
using System;
using System.Threading;
using System.Runtime.CompilerServices;

[assembly: InternalsVisibleTo("PnPFramework.Test")]
namespace PnP.Framework
{
    public class PnPCoreSdk
    {

        private static readonly Lazy<PnPCoreSdk> _lazyInstance = new Lazy<PnPCoreSdk>(() => new PnPCoreSdk(), true);
        private IPnPContextFactory pnpContextFactoryCache;
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

        public PnPContext GetPnPContext(ClientContext context)
        {
            var factory = BuildContextFactory();
            return factory.Create(new Uri(context.Url), AuthenticationProviderFactory.GetAuthenticationProvider(context));
        }

        private IPnPContextFactory BuildContextFactory()
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
                //TODO: Can someone check if this is actually needed or can we just use fluent api? I'm not sure and it may have quite an impact
                services = services.AddPnPCore(options =>
                {
                    options.PnPContext.GraphFirst = false;
                }).Services;
                if(OnDIContainerBuilding != null)
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


    }
}
