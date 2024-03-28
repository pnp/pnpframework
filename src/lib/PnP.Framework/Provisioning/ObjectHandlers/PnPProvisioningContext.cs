using System;
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;

namespace PnP.Framework.Provisioning.ObjectHandlers
{
    /// <summary>
    /// Asynchronous delegate to acquire an Access Token to get access to a target resource
    /// </summary>
    /// <param name="resource">The Resource to access</param>
    /// <param name="scope">The required Permission Scope</param>
    /// <returns>The Access Token to access the target resource</returns>
    public delegate Task<string> AcquireTokenAsyncDelegate(string resource, string scope);

    /// <summary>
    /// Asynchronous delegate to get a cookie to access a target resource
    /// </summary>
    /// <param name="resource">The Resource to access</param>
    /// <returns>The Cookie to access the target resource</returns>
    public delegate Task<string> AcquireCookieAsyncDelegate(string resource);

    /// <summary>
    /// Class to wrap any PnP Provisioning process in order to share the same security context across different Object Handlers
    /// </summary>
    public class PnPProvisioningContext : IDisposable
    {
        private readonly PnPProvisioningContext _previous;

        internal List<string> ParsedSiteUrls { get; private set; } = new List<string>();
        /// <summary>
        /// Asynchronous delegate to acquire an access token for a specific resource and with a specific scope
        /// </summary>
        public AcquireTokenAsyncDelegate AcquireTokenAsync { get; private set; }

        /// <summary>
        /// Asynchronous delegate to acquire a cookie for a specific resource
        /// </summary>
        public AcquireCookieAsyncDelegate AcquireCookieAsync { get; private set; }

        /// <summary>
        /// Property Bag of properties for the current context
        /// </summary>
        public Dictionary<string, object> Properties { get; private set; } =
            new Dictionary<string, object>();

        /// <summary>
        /// Defines the Cloud Deployment the current user is connected to.
        /// </summary>
        public AzureEnvironment AzureEnvironment { get; private set; } = AzureEnvironment.Production;
        /// <summary>
        /// Constructor for the content
        /// </summary>
        /// <param name="acquireTokenAsyncDelegate">Asynchronous delegate to acquire an access token for a specific resource and with a specific scope</param>
        /// <param name="acquireCookieAsyncDelegate">Asynchronous delegate to acquire a cookie for a specific resource</param>
        /// <param name="properties">Properties to add to the Property Bag of the current context</param>
        /// <param name="azureEnvironment">The Azure Cloud Deployment to use. This is used to determine the Graph endpoint. Defaults to 'Production' / 'Global'</param>
        public PnPProvisioningContext(
            AcquireTokenAsyncDelegate acquireTokenAsyncDelegate = null,
            AcquireCookieAsyncDelegate acquireCookieAsyncDelegate = null,
            Dictionary<string, object> properties = null, AzureEnvironment azureEnvironment = AzureEnvironment.Production)
        {
            // Save the delegate to acquire the access token
            this.AcquireTokenAsync = acquireTokenAsyncDelegate;

            // Save the delegate to acquire the cookie
            this.AcquireCookieAsync = acquireCookieAsyncDelegate;

            this.AzureEnvironment = azureEnvironment;

            // Add the initial set of properties, if any
            if (properties != null)
            {
                foreach (var p in properties)
                {
                    this.Properties.Add(p.Key, p.Value);
                }
            }

            // Save the previous context, if any
            this._previous = Current;

            // Set the new context to this
            Current = this;
        }

        /// <summary>
        /// Synchronous delegate to acquire an access token for a specific resource and with a specific scope
        /// </summary>
        /// <param name="resource">The target resource</param>
        /// <param name="scope">The scope for the target resource</param>
        /// <returns>The Access Token for the requested resource, with the requested scope</returns>
        public string AcquireToken(string resource, string scope)
        {
            if (this.AcquireTokenAsync != null)
            {
                return (this.AcquireTokenAsync(resource, scope).GetAwaiter().GetResult());
            }
            else
            {
                return null;
            }
        }

        public string AcquireTokenWithMultipleScopes(string resource, params string[] scope)
        {
            if (this.AcquireTokenAsync != null)
            {
                return this.AcquireTokenAsync(resource, string.Join(" ", scope)).GetAwaiter().GetResult();
            } else {
                return null;
            }
        }

        /// <summary>
        /// Synchronous delegate to acquire a cookie for a specific resource
        /// </summary>
        /// <param name="resource">The target resource</param>
        /// <returns>The Cookie for the requested resource</returns>
        public string AcquireCookie(string resource)
        {
            // If there's a delegate hooked up to the cookie acquiring, trigger it, if not return a null to indicate it's not able to get a token through a cookie
            return this.AcquireCookieAsync != null ? (this.AcquireCookieAsync(resource).GetAwaiter().GetResult()) : null;
        }

        ~PnPProvisioningContext()
        {
            Dispose(false);
        }

        /// <summary>
        /// Performs application-defined tasks associated with freeing, releasing, or resetting unmanaged resources.
        /// </summary>
        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        private void Dispose(bool disposing)
        {
            if (disposing)
            {
                Current = this._previous;
            }
        }

        private static readonly AsyncLocal<PnPProvisioningContext> _pnpSerializationScope = new AsyncLocal<PnPProvisioningContext>();

        public static PnPProvisioningContext Current
        {
            get { return _pnpSerializationScope.Value; }
            set
            {
                _pnpSerializationScope.Value = value;
            }
        }
    }
}
