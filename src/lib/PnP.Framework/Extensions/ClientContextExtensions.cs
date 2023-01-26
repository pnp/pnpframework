using PnP.Framework;
using PnP.Framework.Diagnostics;
using PnP.Framework.Http;
using PnP.Framework.Provisioning.ObjectHandlers;
using PnP.Framework.Sites;
using PnP.Framework.Utilities;
using PnP.Framework.Utilities.Async;
using PnP.Framework.Utilities.Context;
using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Configuration;
using System.Diagnostics;
using System.IdentityModel.Tokens.Jwt;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Reflection;
using System.Text.Json;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Xml;

namespace Microsoft.SharePoint.Client
{
    /// <summary>
    /// Class that deals with cloning client context object, getting access token and validates server version
    /// </summary>
    public static partial class ClientContextExtensions
    {
        private static readonly string userAgentFromConfig = null;

#pragma warning disable CS0169
        private static ConcurrentDictionary<string, (string requestDigest, DateTime expiresOn)> requestDigestInfos = new ConcurrentDictionary<string, (string requestDigest, DateTime expiresOn)>();
#pragma warning restore CS0169

        //private static bool hasAuthCookies;

        /// <summary>
        /// Static constructor, only executed once per class load
        /// </summary>
#pragma warning disable CA1810
        static ClientContextExtensions()
        {
            try
            {
                ClientContextExtensions.userAgentFromConfig = ConfigurationManager.AppSettings["SharePointPnPUserAgent"];
            }
            catch // throws exception if being called from a .NET Standard 2.0 application
            {

            }
            if (string.IsNullOrWhiteSpace(ClientContextExtensions.userAgentFromConfig))
            {
                ClientContextExtensions.userAgentFromConfig = Environment.GetEnvironmentVariable("SharePointPnPUserAgent", EnvironmentVariableTarget.Process);
            }
        }
#pragma warning restore CA1810
        /// <summary>
        /// Clones a ClientContext object while "taking over" the security context of the existing ClientContext instance
        /// </summary>
        /// <param name="clientContext">ClientContext to be cloned</param>
        /// <param name="siteUrl">Site URL to be used for cloned ClientContext</param>
        /// <param name="accessTokens">Dictionary of access tokens for sites URLs</param>
        /// <returns>A ClientContext object created for the passed site URL</returns>
        public static ClientContext Clone(this ClientRuntimeContext clientContext, string siteUrl, Dictionary<string, string> accessTokens = null)
        {
            if (string.IsNullOrWhiteSpace(siteUrl))
            {
                throw new ArgumentException(CoreResources.ClientContextExtensions_Clone_Url_of_the_site_is_required_, nameof(siteUrl));
            }

            return clientContext.Clone(new Uri(siteUrl), accessTokens);
        }

        /// <summary>
        /// Executes the current set of data retrieval queries and method invocations and retries it if needed using the Task Library.
        /// </summary>
        /// <param name="clientContext">clientContext to operate on</param>
        /// <param name="retryCount">Number of times to retry the request</param>
        /// <param name="userAgent">UserAgent string value to insert for this request. You can define this value in your app's config file using key="SharePointPnPUserAgent" value="PnPRocks"></param>
        public static Task ExecuteQueryRetryAsync(this ClientRuntimeContext clientContext, int retryCount = 10, string userAgent = null)
        {
            return ExecuteQueryImplementation(clientContext, retryCount, userAgent);
        }


        /// <summary>
        /// Executes the current set of data retrieval queries and method invocations and retries it if needed.
        /// </summary>
        /// <param name="clientContext">clientContext to operate on</param>
        /// <param name="retryCount">Number of times to retry the request</param>
        /// <param name="userAgent">UserAgent string value to insert for this request. You can define this value in your app's config file using key="SharePointPnPUserAgent" value="PnPRocks"></param>
        public static void ExecuteQueryRetry(this ClientRuntimeContext clientContext, int retryCount = 10, string userAgent = null)
        {
            Task.Run(() => ExecuteQueryImplementation(clientContext, retryCount, userAgent)).GetAwaiter().GetResult();
        }

        private static async Task ExecuteQueryImplementation(ClientRuntimeContext clientContext, int retryCount = 10, string userAgent = null)
        {

            await new SynchronizationContextRemover();

            // Set the TLS preference. Needed on some server os's to work when Office 365 removes support for TLS 1.0
            ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12 | SecurityProtocolType.Tls11 | SecurityProtocolType.Tls;

            var clientTag = string.Empty;
            if (clientContext is PnPClientContext)
            {
                retryCount = (clientContext as PnPClientContext).RetryCount;
                clientTag = (clientContext as PnPClientContext).ClientTag;
            }

            int backoffInterval = 500;
            int retryAttempts = 0;
            int retryAfterInterval = 0;
            bool retry = false;
            ClientRequestWrapper wrapper = null;

            if (retryCount <= 0)
                throw new ArgumentException("Provide a retry count greater than zero.");

            // Do while retry attempt is less than retry count
            while (retryAttempts < retryCount)
            {
                try
                {
                    clientContext.ClientTag = SetClientTag(clientTag);

                    // Make CSOM request more reliable by disabling the return value cache. Given we 
                    // often clone context objects and the default value is
                    clientContext.DisableReturnValueCache = true;
                    // Add event handler to "insert" app decoration header to mark the PnP Sites Core library as a known application
                    EventHandler<WebRequestEventArgs> appDecorationHandler = AttachRequestUserAgent(userAgent);

                    clientContext.ExecutingWebRequest += appDecorationHandler;

                    // DO NOT CHANGE THIS TO EXECUTEQUERYRETRY
                    if (!retry)
                    {
                        await clientContext.ExecuteQueryAsync();
                    }
                    else
                    {
                        if (wrapper != null && wrapper.Value != null)
                        {
                            await clientContext.RetryQueryAsync(wrapper.Value);
                        }
                    }

                    // Remove the app decoration event handler after the executequery
                    clientContext.ExecutingWebRequest -= appDecorationHandler;

                    return;
                }
                catch (WebException wex)
                {
                    var response = wex.Response as HttpWebResponse;
                    // Check if request was throttled - http status code 429
                    // Check is request failed due to server unavailable - http status code 503
                    if ((response != null &&
                        (response.StatusCode == (HttpStatusCode)429
                        || response.StatusCode == (HttpStatusCode)503
                        // || response.StatusCode == (HttpStatusCode)500
                        ))
                        || wex.Status == WebExceptionStatus.Timeout)
                    {
                        wrapper = (ClientRequestWrapper)wex.Data["ClientRequest"];
                        retry = true;
                        retryAfterInterval = 0;

                        //Add delay for retry, retry-after header is specified in seconds
                        if (response != null && response.Headers["Retry-After"] != null)
                        {
                            if (int.TryParse(response.Headers["Retry-After"], out int retryAfterHeaderValue))
                            {
                                retryAfterInterval = retryAfterHeaderValue * 1000;
                            }
                        }
                        else
                        {
                            retryAfterInterval = backoffInterval;
                            backoffInterval *= 2;
                        }

                        if (wex.Status == WebExceptionStatus.Timeout)
                        {
                            Log.Warning(Constants.LOGGING_SOURCE, $"CSOM request timeout. Retry attempt {retryAttempts + 1}. Sleeping for {retryAfterInterval} milliseconds before retrying.");
                        }
                        else
                        {
                            Log.Warning(Constants.LOGGING_SOURCE, $"CSOM request frequency exceeded usage limits. Retry attempt {retryAttempts + 1}. Sleeping for {retryAfterInterval} milliseconds before retrying.");
                        }

                        await Task.Delay(retryAfterInterval);

                        //Add to retry count and increase delay.
                        retryAttempts++;
                    }
                    else
                    {
                        var errorSb = new System.Text.StringBuilder();

                        errorSb.AppendLine(wex.ToString());
                        errorSb.AppendLine($"TraceCorrelationId: {clientContext.TraceCorrelationId}");
                        errorSb.AppendLine($"Url: {clientContext.Url}");

                        //find innermost Error and check if it is a SocketException
                        Exception innermostEx = wex;
                        while (innermostEx.InnerException != null) innermostEx = innermostEx.InnerException;
                        var socketEx = innermostEx as System.Net.Sockets.SocketException;
                        if (socketEx != null)
                        {
                            errorSb.AppendLine($"ErrorCode: {socketEx.ErrorCode}"); //10054
                            errorSb.AppendLine($"SocketErrorCode: {socketEx.SocketErrorCode}"); //ConnectionReset
                            errorSb.AppendLine($"Message: {socketEx.Message}"); //An existing connection was forcibly closed by the remote host
                            Log.Error(Constants.LOGGING_SOURCE, CoreResources.ClientContextExtensions_ExecuteQueryRetryException, errorSb.ToString());

                            // Hostname unknown error code 11001 should not be retried
                            if(socketEx.ErrorCode == 11001)
                            {
                                throw;
                            }

                            //retry
                            wrapper = (ClientRequestWrapper)wex.Data["ClientRequest"];
                            retry = true;
                            retryAfterInterval = 0;

                            //Add delay for retry, retry-after header is specified in seconds
                            if (response != null && response.Headers["Retry-After"] != null)
                            {
                                if (int.TryParse(response.Headers["Retry-After"], out int retryAfterHeaderValue))
                                {
                                    retryAfterInterval = retryAfterHeaderValue * 1000;
                                }
                            }
                            else
                            {
                                retryAfterInterval = backoffInterval;
                                backoffInterval *= 2;
                            }

                            Log.Warning(Constants.LOGGING_SOURCE, $"CSOM request socket exception. Retry attempt {retryAttempts + 1}. Sleeping for {retryAfterInterval} milliseconds before retrying.");

                            await Task.Delay(retryAfterInterval);

                            //Add to retry count and increase delay.
                            retryAttempts++;
                        }
                        else
                        {
                            if (response != null)
                            {
                                //if(response.Headers["SPRequestGuid"] != null) 
                                if (response.Headers.AllKeys.Any(k => string.Equals(k, "SPRequestGuid", StringComparison.InvariantCultureIgnoreCase)))
                                {
                                    var spRequestGuid = response.Headers["SPRequestGuid"];
                                    errorSb.AppendLine($"ServerErrorTraceCorrelationId: {spRequestGuid}");
                                }
                            }

                            Log.Error(Constants.LOGGING_SOURCE, CoreResources.ClientContextExtensions_ExecuteQueryRetryException, errorSb.ToString());
                            throw;
                        }
                    }
                }
                catch (ServerException serverEx)
                {
                    var errorSb = new System.Text.StringBuilder();

                    errorSb.AppendLine(serverEx.ToString());
                    errorSb.AppendLine($"ServerErrorCode: {serverEx.ServerErrorCode}");
                    errorSb.AppendLine($"ServerErrorTypeName: {serverEx.ServerErrorTypeName}");
                    errorSb.AppendLine($"ServerErrorTraceCorrelationId: {serverEx.ServerErrorTraceCorrelationId}");
                    errorSb.AppendLine($"ServerErrorValue: {serverEx.ServerErrorValue}");
                    errorSb.AppendLine($"ServerErrorDetails: {serverEx.ServerErrorDetails}");

                    Log.Error(Constants.LOGGING_SOURCE, CoreResources.ClientContextExtensions_ExecuteQueryRetryException, errorSb.ToString());

                    throw;
                }
            }

            throw new MaximumRetryAttemptedException($"Maximum retry attempts {retryCount}, has be attempted.");
        }

        /// <summary>
        /// Attaches either a passed user agent, or one defined in the App.config file, to the WebRequstExecutor UserAgent property.
        /// </summary>
        /// <param name="customUserAgent">a custom user agent to override any defined in App.config</param>
        /// <returns>An EventHandler of WebRequestEventArgs.</returns>
        private static EventHandler<WebRequestEventArgs> AttachRequestUserAgent(string customUserAgent)
        {
            return (s, e) =>
            {
                bool overrideUserAgent = true;
                var existingUserAgent = e.WebRequestExecutor.WebRequest.UserAgent;
                if (!string.IsNullOrEmpty(existingUserAgent) && existingUserAgent.StartsWith("NONISV|SharePointPnP|PnPPS/"))
                {
                    overrideUserAgent = false;
                }
                if (overrideUserAgent)
                {
                    if (string.IsNullOrEmpty(customUserAgent) && !string.IsNullOrEmpty(ClientContextExtensions.userAgentFromConfig))
                    {
                        customUserAgent = userAgentFromConfig;
                    }
                    e.WebRequestExecutor.WebRequest.UserAgent = string.IsNullOrEmpty(customUserAgent) ? $"{PnPCoreUtilities.PnPCoreUserAgent}" : customUserAgent;
                }
            };
        }

        /// <summary>
        /// Sets the client context client tag on outgoing CSOM requests.
        /// </summary>
        /// <param name="clientTag">An optional client tag to set on client context requests.</param>
        /// <returns></returns>
        private static string SetClientTag(string clientTag = "")
        {
            // ClientTag property is limited to 32 chars
            if (string.IsNullOrEmpty(clientTag))
            {
                clientTag = $"{PnPCoreUtilities.PnPCoreVersionTag}:{GetCallingPnPMethod()}";
            }
            if (clientTag.Length > 32)
            {
                clientTag = clientTag.Substring(0, 32);
            }

            return clientTag;
        }

        /// <summary>
        /// Clones a ClientContext object while "taking over" the security context of the existing ClientContext instance
        /// </summary>
        /// <param name="clientContext">ClientContext to be cloned</param>
        /// <param name="siteUrl">Site URL to be used for cloned ClientContext</param>
        /// <param name="accessTokens">Dictionary of access tokens for sites URLs</param>
        /// <returns>A ClientContext object created for the passed site URL</returns>
        public static ClientContext Clone(this ClientRuntimeContext clientContext, Uri siteUrl, Dictionary<string, string> accessTokens = null)
        {
            return Clone(clientContext, new ClientContext(siteUrl), siteUrl, accessTokens);
        }
        /// <summary>
        /// Clones a ClientContext object while "taking over" the security context of the existing ClientContext instance
        /// </summary>
        /// <param name="clientContext">ClientContext to be cloned</param>
        /// <param name="targetContext">CientContext stub to be used for cloning</param>
        /// <param name="siteUrl">Site URL to be used for cloned ClientContext</param>
        /// <param name="accessTokens">Dictionary of access tokens for sites URLs</param>
        /// <returns>A ClientContext object created for the passed site URL</returns>
        internal static ClientContext Clone(this ClientRuntimeContext clientContext, ClientContext targetContext, Uri siteUrl, Dictionary<string, string> accessTokens = null)
        {
            if (siteUrl == null)
            {
                throw new ArgumentException(CoreResources.ClientContextExtensions_Clone_Url_of_the_site_is_required_, nameof(siteUrl));
            }

            ClientContext clonedClientContext = targetContext;
            clonedClientContext.ClientTag = clientContext.ClientTag;
            clonedClientContext.DisableReturnValueCache = clientContext.DisableReturnValueCache;
            clonedClientContext.WebRequestExecutorFactory = clientContext.WebRequestExecutorFactory;

            // Check if we do have context settings
            var contextSettings = clientContext.GetContextSettings();

            if (contextSettings != null) // We do have more information about this client context, so let's use it to do a more intelligent clone
            {
                string newSiteUrl = siteUrl.ToString();

                // A diffent host = different audience ==> new access token is needed
                if (contextSettings.UsesDifferentAudience(newSiteUrl))
                {

                    var authManager = contextSettings.AuthenticationManager;
                    ClientContext newClientContext = null;
                    if (contextSettings.Type != ClientContextType.Cookie)
                    {
                        if (contextSettings.Type == ClientContextType.SharePointACSAppOnly)
                        {
                            newClientContext = authManager.GetACSAppOnlyContext(newSiteUrl, contextSettings.ClientId, contextSettings.ClientSecret, contextSettings.Environment);
                        }
                        else if (contextSettings.Type == ClientContextType.OnPremises)
                        {
                            newClientContext = authManager.GetOnPremisesContext(newSiteUrl, clientContext.Credentials);
                        }
                        else
                        {
                            newClientContext = authManager.GetContextAsync(newSiteUrl).GetAwaiter().GetResult();
                        }
                    }
                    else
                    {
                        newClientContext = new ClientContext(newSiteUrl);
                        newClientContext.ExecutingWebRequest += (sender, webRequestEventArgs) =>
                        {
                            // Call the ExecutingWebRequest delegate method from the original ClientContext object, but pass along the webRequestEventArgs of 
                            // the new delegate method
                            MethodInfo methodInfo = clientContext.GetType().GetMethod("OnExecutingWebRequest", BindingFlags.Instance | BindingFlags.NonPublic);
                            object[] parametersArray = new object[] { webRequestEventArgs };
                            methodInfo.Invoke(clientContext, parametersArray);
                        };
                        ClientContextSettings clientContextSettings = new ClientContextSettings()
                        {
                            AuthenticationManager = authManager ?? new PnP.Framework.AuthenticationManager(),
                            Type = ClientContextType.Cookie,
                            SiteUrl = newSiteUrl
                        };
                        if (authManager != null)
                        {
                            clientContextSettings.AuthenticationManager.CookieContainer = authManager.CookieContainer;
                        }

                        newClientContext.AddContextSettings(clientContextSettings);
                    }
                    if (newClientContext != null)
                    {
                        //Take over the form digest handling setting
                        newClientContext.ClientTag = clientContext.ClientTag;
                        newClientContext.DisableReturnValueCache = clientContext.DisableReturnValueCache;
                        newClientContext.WebRequestExecutorFactory = clientContext.WebRequestExecutorFactory;
                        return newClientContext;
                    }
                    else
                    {
                        throw new Exception($"Cloning for context setting type {contextSettings.Type} was not yet implemented");
                    }
                }
                else
                {
                    // Take over the context settings, this is needed if we later on want to clone this context to a different audience
                    contextSettings.SiteUrl = newSiteUrl;
                    clonedClientContext.AddContextSettings(contextSettings);

                    if (contextSettings.Type == ClientContextType.OnPremises)
                    {
                        var authManager = contextSettings.AuthenticationManager;
                        clonedClientContext.Credentials = clientContext.Credentials;
                        authManager.ConfigureOnPremisesContext(newSiteUrl, clonedClientContext);
                    }
                    else
                    {
                        clonedClientContext.ExecutingWebRequest += delegate (object oSender, WebRequestEventArgs webRequestEventArgs)
                        {
                            // Call the ExecutingWebRequest delegate method from the original ClientContext object, but pass along the webRequestEventArgs of 
                            // the new delegate method
                            MethodInfo methodInfo = clientContext.GetType().GetMethod("OnExecutingWebRequest", BindingFlags.Instance | BindingFlags.NonPublic);
                            object[] parametersArray = new object[] { webRequestEventArgs };
                            methodInfo.Invoke(clientContext, parametersArray);
                        };
                    }
                }
            }
            else // Fallback the default cloning logic if there were not context settings available
            {
                //Take over the form digest handling setting

                var originalUri = new Uri(clientContext.Url);
                // If the cloned host is not the same as the original one
                // and if there is an active PnPProvisioningContext
                if (originalUri.Host != siteUrl.Host &&
                    PnPProvisioningContext.Current != null)
                {
                    // Let's apply that specific Access Token
                    clonedClientContext.ExecutingWebRequest += (sender, args) =>
                    {
                        // We get a fresh new Access Token for every request, to avoid using an expired one
                        var accessToken = PnPProvisioningContext.Current.AcquireToken(siteUrl.Authority, null);
                        args.WebRequestExecutor.RequestHeaders["Authorization"] = "Bearer " + accessToken;
                    };
                }
                // Else if the cloned host is not the same as the original one
                // and if there is a custom Access Token for it in the input arguments
                else if (originalUri.Host != siteUrl.Host &&
                    accessTokens != null && accessTokens.Count > 0 &&
                    accessTokens.ContainsKey(siteUrl.Authority))
                {
                    // Let's apply that specific Access Token
                    clonedClientContext.ExecutingWebRequest += (sender, args) =>
                    {
                        args.WebRequestExecutor.RequestHeaders["Authorization"] = "Bearer " + accessTokens[siteUrl.Authority];
                    };
                }
                // Else if the cloned host is not the same as the original one
                // and if the client context is a PnPClientContext with custom access tokens in its property bag
                else if (originalUri.Host != siteUrl.Host &&
                    accessTokens == null && clientContext is PnPClientContext &&
                    ((PnPClientContext)clientContext).PropertyBag.ContainsKey("AccessTokens") &&
                    ((Dictionary<string, string>)((PnPClientContext)clientContext).PropertyBag["AccessTokens"]).ContainsKey(siteUrl.Authority))
                {
                    // Let's apply that specific Access Token
                    clonedClientContext.ExecutingWebRequest += (sender, args) =>
                    {
                        args.WebRequestExecutor.RequestHeaders["Authorization"] = "Bearer " + ((Dictionary<string, string>)((PnPClientContext)clientContext).PropertyBag["AccessTokens"])[siteUrl.Authority];
                    };
                }
                else
                {
                    // In case of app only or SAML
                    clonedClientContext.ExecutingWebRequest += (sender, webRequestEventArgs) =>
                    {
                        // Call the ExecutingWebRequest delegate method from the original ClientContext object, but pass along the webRequestEventArgs of 
                        // the new delegate method
                        MethodInfo methodInfo = clientContext.GetType().GetMethod("OnExecutingWebRequest", BindingFlags.Instance | BindingFlags.NonPublic);
                        object[] parametersArray = new object[] { webRequestEventArgs };
                        methodInfo.Invoke(clientContext, parametersArray);
                    };
                }
            }

            return clonedClientContext;
        }

        /// <summary>
        /// Returns the number of pending requests
        /// </summary>
        /// <param name="clientContext">Client context to check the pending requests for</param>
        /// <returns>The number of pending requests</returns>
        public static int PendingRequestCount(this ClientRuntimeContext clientContext)
        {
            int count = 0;

            if (clientContext.HasPendingRequest)
            {
                var result = clientContext.PendingRequest.GetType().GetProperty("Actions", BindingFlags.GetProperty | BindingFlags.Instance | BindingFlags.NonPublic);
                if (result != null)
                {
                    var propValue = result.GetValue(clientContext.PendingRequest);
                    if (propValue != null)
                    {
                        count = (propValue as List<ClientAction>).Count;
                    }
                }
            }

            return count;
        }

        /// <summary>
        /// Gets a site collection context for the passed web. This site collection client context uses the same credentials
        /// as the passed client context
        /// </summary>
        /// <param name="clientContext">Client context to take the credentials from</param>
        /// <returns>A site collection client context object for the site collection</returns>
        public static ClientContext GetSiteCollectionContext(this ClientRuntimeContext clientContext)
        {
            Site site = (clientContext as ClientContext).Site;
            if (!site.IsObjectPropertyInstantiated("Url"))
            {
                clientContext.Load(site);
                clientContext.ExecuteQueryRetry();
            }
            return clientContext.Clone(site.Url);
        }

        /// <summary>
        /// Checks if the used ClientContext is app-only
        /// </summary>
        /// <param name="clientContext">The ClientContext to inspect</param>
        /// <returns>True if app-only, false otherwise</returns>
        public static bool IsAppOnly(this ClientRuntimeContext clientContext)
        {

            // Set initial result to false
            var result = false;

            // do we have cookies?

            // Try to get an access token from the current context
            var accessToken = clientContext.GetAccessToken();

            // If any
            if (!string.IsNullOrEmpty(accessToken))
            {
                // Try to decode the access token
                var token = new JwtSecurityToken(accessToken);

                // Search for the UPN claim, to see if we have user's delegation
                var upn = token.Claims.FirstOrDefault(claim => claim.Type == "upn")?.Value;
                if (string.IsNullOrEmpty(upn))
                {
                    result = true;
                }
            }
            else if (clientContext.Credentials == null)
            {
                result = true;
            }
            if (result == true)
            {
                try
                {
                    var contextSettings = (clientContext as ClientContext).GetContextSettings();
                    if (contextSettings.Type == ClientContextType.Cookie)
                    {
                        result = false;
                    }
                    // var cookieString = CookieReader.GetCookie(clientContext.Url)?.Replace("; ", ",")?.Replace(";", ",");

                    // if (Regex.IsMatch(cookieString, "FedAuth", RegexOptions.IgnoreCase))
                    // {
                    //     result = false;
                    // }
                    // else if (Regex.IsMatch(cookieString, "EdgeAccessCookie", RegexOptions.IgnoreCase))
                    // {
                    //     result = false;
                    // }
                }
                catch (Exception)
                {

                }
            }

            return result;
        }


        /// <summary>
        /// Gets an access token from a <see cref="ClientContext"/> instance. Only works when using an add-in or app-only authentication flow.
        /// </summary>
        /// <param name="clientContext"><see cref="ClientContext"/> instance to obtain an access token for</param>
        /// <returns>Access token for the given <see cref="ClientContext"/> instance</returns>
        public static string GetAccessToken(this ClientRuntimeContext clientContext)
        {
            string accessToken = null;

            if (PnPProvisioningContext.Current?.AcquireTokenAsync != null)
            {
                accessToken = PnPProvisioningContext.Current.AcquireToken(new Uri(clientContext.Url).Authority, null);
            }
            else
            {
                var contextSettings = clientContext.GetContextSettings();

                if (contextSettings?.AuthenticationManager != null && contextSettings?.Type != ClientContextType.SharePointACSAppOnly && contextSettings?.Type != ClientContextType.OnPremises && contextSettings?.Type != ClientContextType.Cookie)
                {
                    accessToken = contextSettings.AuthenticationManager.GetAccessTokenAsync(clientContext.Url).GetAwaiter().GetResult();
                }
                else
                {
                    // Get User Agent String
                    string userAgentFromConfig = null;
                    try
                    {
                        userAgentFromConfig = ConfigurationManager.AppSettings["SharePointPnPUserAgent"];
                    }
                    catch // throws exception if being called from a .NET Standard 2.0 application
                    {

                    }

                    // Get user Agent String if being called from a .NET Standard 2.0 application or is missing
                    if (string.IsNullOrWhiteSpace(userAgentFromConfig))
                    {
                        userAgentFromConfig = Environment.GetEnvironmentVariable("SharePointPnPUserAgent", EnvironmentVariableTarget.Process);
                    }

                    // Use Default User Agent String
                    if (string.IsNullOrWhiteSpace(userAgentFromConfig))
                    {
                        userAgentFromConfig = PnPCoreUtilities.PnPCoreUserAgent;
                    }

                    EventHandler<WebRequestEventArgs> handler = (s, e) =>
                    {
                        string authorization = e.WebRequestExecutor.RequestHeaders["Authorization"];
                        if (!string.IsNullOrEmpty(authorization))
                        {
                            accessToken = authorization.Replace("Bearer ", string.Empty);
                        }

                        e.WebRequestExecutor.WebRequest.UserAgent = string.IsNullOrEmpty(userAgentFromConfig) ? $"{PnPCoreUtilities.PnPCoreUserAgent}" : userAgentFromConfig;
                    };
                    // Issue a dummy request to get it from the Authorization header
                    clientContext.ExecutingWebRequest += handler;
                    clientContext.ExecuteQuery();
                    clientContext.ExecutingWebRequest -= handler;
                }
            }

            return accessToken;
        }

#pragma warning disable CA1034,CA2229,CA1032
        /// <summary>
        /// Defines a Maximum Retry Attemped Exception
        /// </summary>
        [Serializable]
        public class MaximumRetryAttemptedException : Exception
        {
            /// <summary>
            /// Constructor
            /// </summary>
            /// <param name="message"></param>
            public MaximumRetryAttemptedException(string message)
                : base(message)
            {

            }
        }
#pragma warning restore CA1034,CA2229,CA1032

        /// <summary>
        /// Checks the server library version of the context for a minimally required version
        /// </summary>
        /// <param name="clientContext">clientContext to operate on</param>
        /// <param name="minimallyRequiredVersion">provide version to validate</param>
        /// <returns>True if it has minimal required version, false otherwise</returns>
        public static bool HasMinimalServerLibraryVersion(this ClientRuntimeContext clientContext, string minimallyRequiredVersion)
        {
            return HasMinimalServerLibraryVersion(clientContext, new Version(minimallyRequiredVersion));
        }

        /// <summary>
        /// Checks the server library version of the context for a minimally required version
        /// </summary>
        /// <param name="clientContext">clientContext to operate on</param>
        /// <param name="minimallyRequiredVersion">provide version to validate</param>
        /// <returns>True if it has minimal required version, false otherwise</returns>
        public static bool HasMinimalServerLibraryVersion(this ClientRuntimeContext clientContext, Version minimallyRequiredVersion)
        {
            bool hasMinimalVersion = false;
            try
            {
                clientContext.ExecuteQueryRetry();
                hasMinimalVersion = clientContext.ServerLibraryVersion.CompareTo(minimallyRequiredVersion) >= 0;
            }
            catch (PropertyOrFieldNotInitializedException)
            {
                // swallow the exception.
            }

            return hasMinimalVersion;
        }

        /// <summary>
        /// Returns the name of the method calling ExecuteQueryRetry and ExecuteQueryRetryAsync
        /// </summary>
        /// <returns>A string with the method name</returns>
        private static string GetCallingPnPMethod()
        {
            StackTrace t = new StackTrace();

            string pnpMethod = "";
            try
            {
                for (int i = 0; i < t.FrameCount; i++)
                {
                    var frame = t.GetFrame(i);
                    var frameName = frame.GetMethod().Name;
                    if (frameName.Equals("ExecuteQueryRetry") || frameName.Equals("ExecuteQueryRetryAsync"))
                    {
                        var method = t.GetFrame(i + 1).GetMethod();

                        // Only return the calling method in case ExecuteQueryRetry was called from inside the PnP core library
                        if (method.Module.Name.Equals("PnP.Framework.dll", StringComparison.InvariantCultureIgnoreCase))
                        {
                            pnpMethod = method.Name;
                        }
                        break;
                    }
                }
            }
            catch
            {
                // ignored
            }

            return pnpMethod;
        }

        /// <summary>
        /// Returns the request digest from the current session/site given cookie based auth
        /// </summary>
        /// <param name="context"></param>
        /// <param name="cookieContainer">A cookiecontainer containing FedAuth cookies</param>
        /// <returns></returns>
        public static async Task<string> GetRequestDigestAsync(this ClientContext context, CookieContainer cookieContainer)
        {
            if (cookieContainer != null)
            {
                var hostUrl = context.Url;
                if (requestDigestInfos.TryGetValue(hostUrl, out (string digestToken, DateTime expiresOn) requestDigestInfo))
                {
                    // We only have to add a request digest when running in dotnet core
                    if (DateTime.Now > requestDigestInfo.expiresOn)
                    {
                        requestDigestInfo = await GetRequestDigestInfoAsync(hostUrl, cookieContainer);
                        requestDigestInfos.AddOrUpdate(hostUrl, requestDigestInfo, (key, oldValue) => requestDigestInfo);
                    }
                }
                else
                {
                    // admin url maybe?
                    requestDigestInfo = await GetRequestDigestInfoAsync(hostUrl, cookieContainer);
                    requestDigestInfos.AddOrUpdate(hostUrl, requestDigestInfo, (key, oldValue) => requestDigestInfo);
                }
                return requestDigestInfo.digestToken;
            }
            else
            {
                return null;
            }
        }

        private static async Task<(string digestToken, DateTime expiresOn)> GetRequestDigestInfoAsync(string siteUrl, CookieContainer cookieContainer)
        {
            await new SynchronizationContextRemover();

            var httpClient = PnPHttpClient.Instance.GetHttpClient();

            string requestUrl = string.Format("{0}/_api/contextinfo", siteUrl.TrimEnd('/'));
            using (HttpRequestMessage request = new HttpRequestMessage(HttpMethod.Post, requestUrl))
            {
                request.Headers.Add("accept", "application/json;odata=nometadata");

                request.Headers.Add("Cookie", cookieContainer.GetCookieHeader(new Uri(siteUrl)));

                HttpResponseMessage response = await httpClient.SendAsync(request);


                string responseString;
                if (response.IsSuccessStatusCode)
                {
                    responseString = await response.Content.ReadAsStringAsync();
                }
                else
                {
                    var errorSb = new System.Text.StringBuilder();

                    errorSb.AppendLine(await response.Content.ReadAsStringAsync());
                    if (response.Headers.Contains("SPRequestGuid"))
                    {
                        var values = response.Headers.GetValues("SPRequestGuid");
                        if (values != null)
                        {
                            var spRequestGuid = values.FirstOrDefault();
                            errorSb.AppendLine($"ServerErrorTraceCorrelationId: {spRequestGuid}");
                        }
                    }

                    throw new Exception(errorSb.ToString());
                }

                var contextInformation = JsonSerializer.Deserialize<JsonElement>(responseString);

                string formDigestValue = contextInformation.GetProperty("FormDigestValue").GetString();
                int expiresIn = contextInformation.GetProperty("FormDigestTimeoutSeconds").GetInt32();
                return (formDigestValue, DateTime.Now.AddSeconds(expiresIn - 30));
            }
        }

        public static async Task<string> GetRequestDigestAsync(this ClientContext context)
        {
            var hostUrl = context.Url;
            if (requestDigestInfos.TryGetValue(hostUrl, out (string digestToken, DateTime expiresOn) requestDigestInfo))
            {
                // We only have to add a request digest when running in dotnet core
                if (DateTime.Now > requestDigestInfo.expiresOn)
                {
                    requestDigestInfo = await GetRequestDigestInfoAsync(context);
                    requestDigestInfos.AddOrUpdate(hostUrl, requestDigestInfo, (key, oldValue) => requestDigestInfo);
                }
            }
            else
            {
                // admin url maybe?
                requestDigestInfo = await GetRequestDigestInfoAsync(context);
                requestDigestInfos.AddOrUpdate(hostUrl, requestDigestInfo, (key, oldValue) => requestDigestInfo);
            }
            return requestDigestInfo.digestToken;
        }
        /// <summary>
        /// Returns the request digest from the current session/site
        /// </summary>
        /// <param name="context"></param>
        /// <returns></returns>
        private static async Task<(string digestToken, DateTime expiresOn)> GetRequestDigestInfoAsync(ClientContext context)
        {
            await new SynchronizationContextRemover();

            string responseString = string.Empty;
            var accessToken = context.GetAccessToken();

            context.Web.EnsureProperty(w => w.Url);

            var httpClient = PnPHttpClient.Instance.GetHttpClient();

            string requestUrl = String.Format("{0}/_api/contextinfo", context.Url);
            using (var request = new HttpRequestMessage(HttpMethod.Post, requestUrl))
            {
                request.Headers.Add("accept", "application/json;odata=nometadata");
                if (!string.IsNullOrEmpty(accessToken))
                {
                    request.Headers.Authorization = new System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", accessToken);
                }

                HttpResponseMessage response = await httpClient.SendAsync(request);

                if (response.IsSuccessStatusCode)
                {
                    responseString = await response.Content.ReadAsStringAsync();
                }
                else
                {
                    var errorSb = new System.Text.StringBuilder();

                    errorSb.AppendLine(await response.Content.ReadAsStringAsync());
                    if (response.Headers.Contains("SPRequestGuid"))
                    {
                        var values = response.Headers.GetValues("SPRequestGuid");
                        if (values != null)
                        {
                            var spRequestGuid = values.FirstOrDefault();
                            errorSb.AppendLine($"ServerErrorTraceCorrelationId: {spRequestGuid}");
                        }
                    }

                    throw new Exception(errorSb.ToString());
                }
            }
            var contextInformation = JsonSerializer.Deserialize<JsonElement>(responseString);

            string formDigestValue = contextInformation.GetProperty("FormDigestValue").GetString();
            int expiresIn = contextInformation.GetProperty("FormDigestTimeoutSeconds").GetInt32();
            return (formDigestValue, DateTime.Now.AddSeconds(expiresIn - 30));
        }

        internal static async Task<string> GetOnPremisesRequestDigestAsync(this ClientContext context)
        {
            var hostUrl = context.Url;
            if (requestDigestInfos.TryGetValue(hostUrl, out (string digestToken, DateTime expiresOn) requestDigestInfo))
            {
                // We only have to add a request digest when running in dotnet core
                if (DateTime.Now > requestDigestInfo.expiresOn)
                {
                    requestDigestInfo = await GetOnPremisesRequestDigestInfoAsync(context);
                    requestDigestInfos.AddOrUpdate(hostUrl, requestDigestInfo, (key, oldValue) => requestDigestInfo);
                }
            }
            else
            {
                // admin url maybe?
                requestDigestInfo = await GetOnPremisesRequestDigestInfoAsync(context);
                requestDigestInfos.AddOrUpdate(hostUrl, requestDigestInfo, (key, oldValue) => requestDigestInfo);
            }
            return requestDigestInfo.digestToken;
        }

        private static async Task<(string digestToken, DateTime expiresOn)> GetOnPremisesRequestDigestInfoAsync(ClientContext context)
        {
            await new SynchronizationContextRemover();

            string responseString = string.Empty;

            string requestUrl = $"{context.Url}/_vti_bin/sites.asmx";

            StringContent content = new StringContent("<?xml version=\"1.0\" encoding=\"utf-8\"?><soap:Envelope xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\" xmlns:xsd=\"http://www.w3.org/2001/XMLSchema\" xmlns:soap=\"http://schemas.xmlsoap.org/soap/envelope/\"><soap:Body><GetUpdatedFormDigestInformation xmlns=\"http://schemas.microsoft.com/sharepoint/soap/\" /></soap:Body></soap:Envelope>");
            // Remove the default Content-Type content header
            if (content.Headers.Contains("Content-Type"))
            {
                content.Headers.Remove("Content-Type");
            }
            // Add the batch Content-Type header
            content.Headers.Add($"Content-Type", "text/xml");
            content.Headers.Add("SOAPAction", "http://schemas.microsoft.com/sharepoint/soap/GetUpdatedFormDigestInformation");
            content.Headers.Add("X-RequestForceAuthentication", "true");

            using (var request = new HttpRequestMessage(HttpMethod.Post, requestUrl))
            {
                request.Content = content;

#pragma warning disable CA2000 // Dispose objects before losing scope
                var httpClient = PnPHttpClient.Instance.GetHttpClient(context);
#pragma warning restore CA2000 // Dispose objects before losing scope

                //Note: no credentials are passed here because the returned http context uses an already correctly configured handler

                HttpResponseMessage response = await httpClient.SendAsync(request);

                if (response.IsSuccessStatusCode)
                {
                    responseString = await response.Content.ReadAsStringAsync();
                }
                else
                {
                    var errorSb = new System.Text.StringBuilder();

                    errorSb.AppendLine(await response.Content.ReadAsStringAsync());
                    if (response.Headers.Contains("SPRequestGuid"))
                    {
                        var values = response.Headers.GetValues("SPRequestGuid");
                        if (values != null)
                        {
                            var spRequestGuid = values.FirstOrDefault();
                            errorSb.AppendLine($"ServerErrorTraceCorrelationId: {spRequestGuid}");
                        }
                    }

                    throw new Exception(errorSb.ToString());
                }
            }

            XmlDocument xd = new XmlDocument();
            xd.LoadXml(responseString);

            XmlNamespaceManager nsmgr = new XmlNamespaceManager(xd.NameTable);
            nsmgr.AddNamespace("soap", "http://schemas.microsoft.com/sharepoint/soap/");
            XmlNode digestNode = xd.SelectSingleNode("//soap:DigestValue", nsmgr);
            if (digestNode != null)
            {
                XmlNode timeOutNode = xd.SelectSingleNode("//soap:TimeoutSeconds", nsmgr);
                int expiresIn = int.Parse(timeOutNode.InnerText);
                return (digestNode.InnerText, DateTime.Now.AddSeconds(expiresIn - 30));
            }
            else
            {
                throw new Exception("No digest found!");
            }
        }

        /// <summary>
        /// BETA: Creates a Communication Site Collection
        /// </summary>
        /// <param name="clientContext"></param>
        /// <param name="siteCollectionCreationInformation"></param>
        /// <returns></returns>
        public static async Task<ClientContext> CreateSiteAsync(this ClientContext clientContext, CommunicationSiteCollectionCreationInformation siteCollectionCreationInformation)
        {
            await new SynchronizationContextRemover();

            return await SiteCollection.CreateAsync(clientContext, siteCollectionCreationInformation);
        }

        /// <summary>
        /// BETA: Creates a Team Site Collection with no group
        /// </summary>
        /// <param name="clientContext"></param>
        /// <param name="siteCollectionCreationInformation"></param>
        /// <returns></returns>
        public static async Task<ClientContext> CreateSiteAsync(this ClientContext clientContext, TeamNoGroupSiteCollectionCreationInformation siteCollectionCreationInformation)
        {
            await new SynchronizationContextRemover();

            return await SiteCollection.CreateAsync(clientContext, siteCollectionCreationInformation);
        }

        /// <summary>
        /// BETA: Creates a Team Site Collection
        /// </summary>
        /// <param name="clientContext"></param>
        /// <param name="siteCollectionCreationInformation"></param>
        /// <returns></returns>
        public static async Task<ClientContext> CreateSiteAsync(this ClientContext clientContext, TeamSiteCollectionCreationInformation siteCollectionCreationInformation)
        {
            await new SynchronizationContextRemover();

            return await SiteCollection.CreateAsync(clientContext, siteCollectionCreationInformation);
        }

        /// <summary>
        /// BETA: Groupifies a classic Team Site Collection
        /// </summary>
        /// <param name="clientContext">ClientContext instance of the site to be groupified</param>
        /// <param name="siteCollectionGroupifyInformation">Information needed to groupify this site</param>
        /// <returns>The clientcontext of the groupified site</returns>
        public static async Task<ClientContext> GroupifySiteAsync(this ClientContext clientContext, TeamSiteCollectionGroupifyInformation siteCollectionGroupifyInformation)
        {
            await new SynchronizationContextRemover();

            return await SiteCollection.GroupifyAsync(clientContext, siteCollectionGroupifyInformation);
        }

        /// <summary>
        /// Checks if an alias is already used for an office 365 group or not
        /// </summary>
        /// <param name="clientContext">ClientContext of the site to operate against</param>
        /// <param name="alias">Alias to verify</param>
        /// <returns>True if in use, false otherwise</returns>
        public static async Task<bool> AliasExistsAsync(this ClientContext clientContext, string alias)
        {
            await new SynchronizationContextRemover();

            return await SiteCollection.AliasExistsAsync(clientContext, alias);
        }

        /// <summary>
        /// Enable MS Teams team on a group connected team site
        /// </summary>
        /// <param name="clientContext"></param>
        /// <param name="graphAccessToken"></param>
        /// <returns></returns>
        public static async Task<string> TeamifyAsync(this ClientContext clientContext, string graphAccessToken = null)
        {
            await new SynchronizationContextRemover();

            return await SiteCollection.TeamifySiteAsync(clientContext, graphAccessToken);
        }


        /// <summary>
        /// Checks whether the teamify prompt is hidden in O365 Group connected sites
        /// </summary>
        /// <param name="clientContext">ClientContext of the site to operate against</param>
        /// <returns></returns>
        public static async Task<bool> IsTeamifyPromptHiddenAsync(this ClientContext clientContext)
        {
            await new SynchronizationContextRemover();

            return await SiteCollection.IsTeamifyPromptHiddenAsync(clientContext);
        }

        [Obsolete("Use IsTeamifyPromptHiddenAsync")]
        public static async Task<bool> IsTeamifyPromptHidden(this ClientContext clientContext)
        {
            return await IsTeamifyPromptHiddenAsync(clientContext);
        }

        /// <summary>
        /// Hide the teamify prompt displayed in O365 group connected sites
        /// </summary>
        /// <param name="clientContext">ClientContext of the site to operate against</param>
        /// <returns></returns>
        public static async Task<bool> HideTeamifyPromptAsync(this ClientContext clientContext)
        {
            await new SynchronizationContextRemover();
            return await SiteCollection.HideTeamifyPromptAsync(clientContext);
        }

        /// <summary>
        /// Deletes a Communication site or a group-less Modern team site
        /// </summary>
        /// <param name="clientContext"></param>
        /// <returns></returns>
        public static async Task<bool> DeleteSiteAsync(this ClientContext clientContext)
        {
            await new SynchronizationContextRemover();

            return await SiteCollection.DeleteSiteAsync(clientContext);
        }

        internal static CookieContainer GetAuthenticationCookies(this ClientContext context)
        {
            var authCookiesContainer = context.GetContextSettings()?.AuthenticationManager.CookieContainer;
            if (authCookiesContainer == null)
            {
                var cookieString = CookieReader.GetCookie(context.Url)?.Replace("; ", ",")?.Replace(";", ",");
                if (cookieString == null)
                {
                    return null;
                }
                authCookiesContainer = new CookieContainer();
                // Get FedAuth and rtFa cookies issued by ADFS when accessing claims aware applications.
                // - or get the EdgeAccessCookie issued by the Web Application Proxy (WAP) when accessing non-claims aware applications (Kerberos).
                IEnumerable<string> authCookies = null;
                if (Regex.IsMatch(cookieString, "FedAuth", RegexOptions.IgnoreCase))
                {
                    authCookies = cookieString.Split(',').Where(c => c.StartsWith("FedAuth", StringComparison.InvariantCultureIgnoreCase) || c.StartsWith("rtFa", StringComparison.InvariantCultureIgnoreCase));
                }
                else if (Regex.IsMatch(cookieString, "EdgeAccessCookie", RegexOptions.IgnoreCase))
                {
                    authCookies = cookieString.Split(',').Where(c => c.StartsWith("EdgeAccessCookie", StringComparison.InvariantCultureIgnoreCase));
                }
                if (authCookies != null)
                {
                    var siteUri = new Uri(context.Url);
                    var extension = siteUri.Host.Substring(siteUri.Host.LastIndexOf('.') + 1);
                    var cookieCollection = new CookieCollection();
                    foreach (var cookie in authCookies)
                    {
                        var cookieName = cookie.Substring(0, cookie.IndexOf("=")); // cannot use split as there might '=' in the value
                        var cookieValue = cookie.Substring(cookieName.Length + 1);
                        cookieCollection.Add(new Cookie(cookieName, cookieValue));
                    }
                    authCookiesContainer.Add(new Uri($"{siteUri.Scheme}://{siteUri.Host}"), cookieCollection);
                    var adminSiteUri = new Uri(siteUri.Scheme + "://" + siteUri.Authority.Replace($".sharepoint.{extension}", $"-admin.sharepoint.{extension}"));
                    authCookiesContainer.Add(adminSiteUri, cookieCollection);
                }
            }
            return authCookiesContainer;
        }
    }
}
