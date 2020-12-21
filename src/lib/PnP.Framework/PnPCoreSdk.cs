using Microsoft.Extensions.DependencyInjection;
using Microsoft.SharePoint.Client;
using PnP.Core.Services;
using System;
using System.Threading;

namespace PnP.Framework
{
    public class PnPCoreSdk
    {

        private static readonly Lazy<PnPCoreSdk> _lazyInstance = new Lazy<PnPCoreSdk>(() => new PnPCoreSdk(), true);
        private IPnPContextFactory pnpContextFactoryCache;
        private static readonly SemaphoreSlim semaphoreSlimFactory = new SemaphoreSlim(1);

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
            return factory.Create(new Uri(context.Url), new PnPCoreSdkAuthenticationProvider(context));
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
                var serviceProvider = new ServiceCollection()
                    .AddPnPCore(options =>
                    {
                        options.PnPContext.GraphFirst = false;
                    })
                    .Services
                .BuildServiceProvider();
                
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
