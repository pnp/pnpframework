using Microsoft.IdentityModel.Clients.ActiveDirectory;
using Microsoft.SharePoint.Client;
using Newtonsoft.Json.Linq;
using PnP.Framework.Diagnostics;
using PnP.Framework.Utilities.Async;
using PnP.Framework.Utilities.Context;
using System;
using System.Globalization;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Runtime.CompilerServices;
using System.Security;
using System.Security.Cryptography.X509Certificates;
using System.Threading;
using System.Threading.Tasks;
using System.Web;

#if DEBUG
[assembly: InternalsVisibleTo("PnP.Framework.Test")]
#endif
namespace PnP.Framework
{
    /// <summary>
    /// Enum to identify the supported Office 365 hosting environments
    /// </summary>
    public enum AzureEnvironment
    {
        Production = 0,
        PPE = 1,
        China = 2,
        Germany = 3,
        USGovernment = 4
    }

    /// <summary>
    /// This manager class can be used to obtain a SharePointContext object
    /// </summary>
    ///

    public enum KnownClientId
    {
        PnPManagementShell,
        SPOManagementShell
    }

    public class AuthenticationManager : IDisposable
    {
        private const string SHAREPOINT_PRINCIPAL = "00000003-0000-0ff1-ce00-000000000000";
        public const string CLIENTID_SPOMANAGEMENTSHELL = "9bc3ab49-b65d-410a-85ad-de819febfddc";
        public const string CLIENTID_PNPMANAGEMENTSHELL = "31359c7f-bd7e-475c-86db-fdb8c937548e";

        private string appOnlyAccessToken;
        private AutoResetEvent appOnlyAccessTokenResetEvent = null;
        private string azureADCredentialsToken;
        private AutoResetEvent azureADCredentialsResetEvent = null;
        private readonly object tokenLock = new object();
        private string _contextUrl;
        private TokenCache _tokenCache;
        private string _commonAuthority = "https://login.windows.net/Common";
        private static AuthenticationContext _authContext = null;
        private string _clientId;
        private Uri _redirectUri;
        private bool disposedValue;

        #region Construction
        public AuthenticationManager()
        {
            // Set the TLS preference. Needed on some server os's to work when Office 365 removes support for TLS 1.0
            ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12 | SecurityProtocolType.Tls11 | SecurityProtocolType.Tls;
        }
        #endregion

        private static string GetKnownClientId(KnownClientId id)
        {
            switch (id)
            {
                case KnownClientId.PnPManagementShell:
                    {
                        return CLIENTID_PNPMANAGEMENTSHELL;
                    }
                case KnownClientId.SPOManagementShell:
                    {
                        return CLIENTID_SPOMANAGEMENTSHELL;
                    }
                default:
                    {
                        return CLIENTID_SPOMANAGEMENTSHELL;
                    }
            }
        }

        /// <summary>
        /// Returns an app only ClientContext object
        /// </summary>
        /// <param name="siteUrl">Site for which the ClientContext object will be instantiated</param>
        /// <param name="appId">Application ID which is requesting the ClientContext object</param>
        /// <param name="appSecret">Application secret of the Application which is requesting the ClientContext object</param>
        /// <returns>ClientContext to be used by CSOM code</returns>
        public ClientContext GetAppOnlyAuthenticatedContext(string siteUrl, string appId, string appSecret)
        {
            return GetAppOnlyAuthenticatedContext(siteUrl, Utilities.TokenHelper.GetRealmFromTargetUrl(new Uri(siteUrl)), appId, appSecret);
        }

        /// <summary>
        /// Returns an app only ClientContext object
        /// </summary>
        /// <param name="siteUrl">Site for which the ClientContext object will be instantiated</param>
        /// <param name="appId">Application ID which is requesting the ClientContext object</param>
        /// <param name="appSecret">Application secret of the Application which is requesting the ClientContext object</param>
        /// <param name="environment">SharePoint environment being used</param>
        /// <returns>ClientContext to be used by CSOM code</returns>
        public ClientContext GetAppOnlyAuthenticatedContext(string siteUrl, string appId, string appSecret, AzureEnvironment environment = AzureEnvironment.Production)
        {
            return GetAppOnlyAuthenticatedContext(siteUrl, Utilities.TokenHelper.GetRealmFromTargetUrl(new Uri(siteUrl)), appId, appSecret, GetAzureADACSEndPoint(environment), GetAzureADACSEndPointPrefix(environment));
        }

        /// <summary>
        /// Returns an app only ClientContext object
        /// </summary>
        /// <param name="siteUrl">Site for which the ClientContext object will be instantiated</param>
        /// <param name="realm">Realm of the environment (tenant) that requests the ClientContext object</param>
        /// <param name="appId">Application ID which is requesting the ClientContext object</param>
        /// <param name="appSecret">Application secret of the Application which is requesting the ClientContext object</param>
        /// <param name="acsHostUrl">Azure ACS host, defaults to accesscontrol.windows.net but internal pre-production environments use other hosts</param>
        /// <param name="globalEndPointPrefix">Azure ACS endpoint prefix, defaults to accounts but internal pre-production environments use other prefixes</param>
        /// <returns>ClientContext to be used by CSOM code</returns>
        public ClientContext GetAppOnlyAuthenticatedContext(string siteUrl, string realm, string appId, string appSecret, string acsHostUrl = "accesscontrol.windows.net", string globalEndPointPrefix = "accounts")
        {
            EnsureToken(siteUrl, realm, appId, appSecret, acsHostUrl, globalEndPointPrefix);
            ClientContext clientContext = Utilities.TokenHelper.GetClientContextWithAccessToken(siteUrl, appOnlyAccessToken);
            clientContext.DisableReturnValueCache = true;

            ClientContextSettings clientContextSettings = new ClientContextSettings()
            {
                Type = ClientContextType.SharePointACSAppOnly,
                SiteUrl = siteUrl,
                AuthenticationManager = this,
                Realm = realm,
                ClientId = appId,
                ClientSecret = appSecret,
                AcsHostUrl = acsHostUrl,
                GlobalEndPointPrefix = globalEndPointPrefix
            };

            clientContext.AddContextSettings(clientContextSettings);

            return clientContext;
        }

        /// <summary>
        /// Get's the Azure ASC login end point for the given environment
        /// </summary>
        /// <param name="environment">Environment to get the login information for</param>
        /// <returns>Azure ASC login endpoint</returns>
        public string GetAzureADACSEndPoint(AzureEnvironment environment)
        {
            switch (environment)
            {
                case AzureEnvironment.Production:
                    {
                        return "accesscontrol.windows.net";
                    }
                case AzureEnvironment.Germany:
                    {
                        return "microsoftonline.de";
                    }
                case AzureEnvironment.China:
                    {
                        return "accesscontrol.chinacloudapi.cn";
                    }
                case AzureEnvironment.USGovernment:
                    {
                        return "microsoftonline.us";
                    }
                case AzureEnvironment.PPE:
                    {
                        return "windows-ppe.net";
                    }
                default:
                    {
                        return "accesscontrol.windows.net";
                    }
            }
        }

        /// <summary>
        /// Get's the Azure ACS login end point prefix for the given environment
        /// </summary>
        /// <param name="environment">Environment to get the login information for</param>
        /// <returns>Azure ACS login endpoint prefix</returns>
        public string GetAzureADACSEndPointPrefix(AzureEnvironment environment)
        {
            switch (environment)
            {
                case AzureEnvironment.Production:
                    {
                        return "accounts";
                    }
                case AzureEnvironment.Germany:
                    {
                        return "login";
                    }
                case AzureEnvironment.China:
                    {
                        return "accounts";
                    }
                case AzureEnvironment.USGovernment:
                    {
                        return "login";
                    }
                case AzureEnvironment.PPE:
                    {
                        return "login";
                    }
                default:
                    {
                        return "accounts";
                    }
            }
        }

        /// <summary>
        /// Ensure that AppAccessToken is filled with a valid string representation of the OAuth AccessToken. This method will launch handle with token cleanup after the token expires
        /// </summary>
        /// <param name="siteUrl">Site for which the ClientContext object will be instantiated</param>
        /// <param name="realm">Realm of the environment (tenant) that requests the ClientContext object</param>
        /// <param name="appId">Application ID which is requesting the ClientContext object</param>
        /// <param name="appSecret">Application secret of the Application which is requesting the ClientContext object</param>
        /// <param name="acsHostUrl">Azure ACS host, defaults to accesscontrol.windows.net but internal pre-production environments use other hosts</param>
        /// <param name="globalEndPointPrefix">Azure ACS endpoint prefix, defaults to accounts but internal pre-production environments use other prefixes</param>
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Globalization", "CA1303:Do not pass literals as localized parameters", MessageId = "OfficeDevPnP.Core.Diagnostics.Log.Debug(System.String,System.String,System.Object[])")]
        private void EnsureToken(string siteUrl, string realm, string appId, string appSecret, string acsHostUrl, string globalEndPointPrefix)
        {
            if (appOnlyAccessToken == null)
            {
                lock (tokenLock)
                {
                    Log.Debug(Constants.LOGGING_SOURCE, "AuthenticationManager:EnsureToken(siteUrl:{0},realm:{1},appId:{2},appSecret:PRIVATE)", siteUrl, realm, appId);
                    if (appOnlyAccessToken == null)
                    {
                        Utilities.TokenHelper.Realm = realm;
                        Utilities.TokenHelper.ServiceNamespace = realm;
                        Utilities.TokenHelper.ClientId = appId;
                        Utilities.TokenHelper.ClientSecret = appSecret;

                        if (!String.IsNullOrEmpty(acsHostUrl))
                        {
                            Utilities.TokenHelper.AcsHostUrl = acsHostUrl;
                        }

                        if (globalEndPointPrefix != null)
                        {
                            Utilities.TokenHelper.GlobalEndPointPrefix = globalEndPointPrefix;
                        }

                        var response = Utilities.TokenHelper.GetAppOnlyAccessToken(SHAREPOINT_PRINCIPAL, new Uri(siteUrl).Authority, realm);
                        string token = response.AccessToken;

                        try
                        {
                            Log.Debug(Constants.LOGGING_SOURCE, "Lease expiration date: {0}", response.ExpiresOn);
                            var lease = GetAccessTokenLease(response.ExpiresOn);
                            lease = TimeSpan.FromSeconds(lease.TotalSeconds - TimeSpan.FromMinutes(5).TotalSeconds > 0 ? lease.TotalSeconds - TimeSpan.FromMinutes(5).TotalSeconds : lease.TotalSeconds);



                            appOnlyAccessTokenResetEvent = new AutoResetEvent(false);

                            AppOnlyAccessTokenWaitInfo wi = new AppOnlyAccessTokenWaitInfo();

                            wi.Handle = ThreadPool.RegisterWaitForSingleObject(appOnlyAccessTokenResetEvent,
                                                                               new WaitOrTimerCallback(AppOnlyAccessTokenWaitProc),
                                                                               wi,
                                                                               (uint)lease.TotalMilliseconds,
                                                                               true);
                        }
                        catch (Exception ex)
                        {
                            Log.Warning(Constants.LOGGING_SOURCE, CoreResources.AuthenticationManger_ProblemDeterminingTokenLease, ex);
                            appOnlyAccessToken = null;
                        }

                        //ThreadPool.QueueUserWorkItem(obj =>
                        //{
                        //    try
                        //    {
                        //        Log.Debug(Constants.LOGGING_SOURCE, "Lease expiration date: {0}", response.ExpiresOn);
                        //        var lease = GetAccessTokenLease(response.ExpiresOn);
                        //        lease =
                        //            TimeSpan.FromSeconds(lease.TotalSeconds - TimeSpan.FromMinutes(5).TotalSeconds > 0 ? lease.TotalSeconds - TimeSpan.FromMinutes(5).TotalSeconds : lease.TotalSeconds);
                        //        Thread.Sleep(lease);
                        //        appOnlyAccessToken = null;
                        //    }
                        //    catch (Exception ex)
                        //    {
                        //        Log.Warning(Constants.LOGGING_SOURCE, CoreResources.AuthenticationManger_ProblemDeterminingTokenLease, ex);
                        //        appOnlyAccessToken = null;
                        //    }
                        //});

                        appOnlyAccessToken = token;
                    }
                }
            }
        }

        internal class AppOnlyAccessTokenWaitInfo
        {
            public RegisteredWaitHandle Handle = null;
        }

        internal void AppOnlyAccessTokenWaitProc(object state, bool timedOut)
        {
            if (!timedOut)
            {
                AppOnlyAccessTokenWaitInfo wi = (AppOnlyAccessTokenWaitInfo)state;
                if (wi.Handle != null)
                {
                    wi.Handle.Unregister(null);
                }
            }
            else
            {
                appOnlyAccessToken = null;
            }
        }



        /// <summary>
        /// Get the access token lease time span.
        /// </summary>
        /// <param name="expiresOn">The ExpiresOn time of the current access token</param>
        /// <returns>Returns a TimeSpan represents the time interval within which the current access token is valid thru.</returns>
        private TimeSpan GetAccessTokenLease(DateTime expiresOn)
        {
            DateTime now = DateTime.UtcNow;
            DateTime expires = expiresOn.Kind == DateTimeKind.Utc ?
                expiresOn : TimeZoneInfo.ConvertTimeToUtc(expiresOn);
            TimeSpan lease = expires - now;
            return lease;
        }

        /// <summary>
        /// Returns a SharePoint ClientContext using Azure Active Directory credential authentication. This depends on the SPO Management Shell app or the PnP Management Shell app being registered in your Azure AD.
        /// </summary>
        /// <param name="siteUrl">Site for which the ClientContext object will be instantiated</param>
        /// <param name="userPrincipalName">The user id</param>
        /// <param name="userPassword">The user's password as a secure string</param>
        /// <param name="environment">SharePoint environment being used</param>
        /// <param name="clientId">Enum value pointing to one of the known client ids</param></parm>
        /// <returns>Client context object</returns>
        public ClientContext GetAzureADCredentialsContext(string siteUrl, string userPrincipalName, SecureString userPassword, AzureEnvironment environment = AzureEnvironment.Production, KnownClientId clientId = KnownClientId.SPOManagementShell)
        {
            string password = new System.Net.NetworkCredential(string.Empty, userPassword).Password;
            return GetAzureADCredentialsContext(siteUrl, userPrincipalName, password, environment, clientId);
        }

        /// <summary>
        /// Returns a SharePoint ClientContext using Azure Active Directory credential authentication. This depends on the SPO Management Shell app or the PnP Management Shell app being registered in your Azure AD.
        /// </summary>
        /// <param name="siteUrl">Site for which the ClientContext object will be instantiated</param>
        /// <param name="userPrincipalName">The user id</param>
        /// <param name="userPassword">The user's password as a string</param>
        /// <param name="environment">SharePoint environment being used</param>
        /// <returns>Client context object</returns>
        public ClientContext GetAzureADCredentialsContext(string siteUrl, string userPrincipalName, string userPassword, AzureEnvironment environment = AzureEnvironment.Production, KnownClientId clientId = KnownClientId.SPOManagementShell)
        {
            Log.Info(Constants.LOGGING_SOURCE, CoreResources.AuthenticationManager_GetContext, siteUrl);
            Log.Debug(Constants.LOGGING_SOURCE, CoreResources.AuthenticationManager_TenantUser, userPrincipalName);

            var spUri = new Uri(siteUrl);
            string resourceUri = spUri.Scheme + "://" + spUri.Authority;

            var clientContext = new ClientContext(siteUrl)
            {
                DisableReturnValueCache = true
            };
            clientContext.ExecutingWebRequest += (sender, args) =>
            {
                EnsureAzureADCredentialsToken(resourceUri, userPrincipalName, userPassword, environment, clientId);
                args.WebRequestExecutor.RequestHeaders["Authorization"] = "Bearer " + azureADCredentialsToken;
            };

            ClientContextSettings clientContextSettings = new ClientContextSettings()
            {
                Type = ClientContextType.AzureADCredentials,
                SiteUrl = siteUrl,
                AuthenticationManager = this,
                UserName = userPrincipalName,
                Password = userPassword
            };

            clientContext.AddContextSettings(clientContextSettings);

            return clientContext;
        }

        /// <summary>
        /// Acquires an access token using Azure AD credential flow. This depends on the SPO Management Shell app or the PnP Management Shell app  being registered in your Azure AD.
        /// </summary>
        /// <param name="resourceUri">Resouce to request access for</param>
        /// <param name="username">User id</param>
        /// <param name="password">Password</param>
        /// <param name="environment">SharePoint environment being used</param>
        /// <param name="clientId">Defaults to the SPO Management Shell client id. Alternatively provide the CLIENTID_PNPMANAGEMENTSHELL or your own client with appropriate permission scopes configured.</param>
        /// <returns>Acces token</returns>
        public static async Task<string> AcquireTokenAsync(string resourceUri, string username, string password, AzureEnvironment environment, string clientId = null)
        {
            return await AcquireTokenAsync(resourceUri, username, password, environment, clientId, null);
        }

        public static async Task<string> AcquireTokenAsync(string resourceUri, string username, string password, AzureEnvironment environment, string clientId = null, Action<string> errorCallback = null)
        {
            HttpClient client = new HttpClient();
            string tokenEndpoint = $"{new AuthenticationManager().GetAzureADLoginEndPoint(environment)}/common/oauth2/token";

            if (clientId == null)
            {
                clientId = GetKnownClientId(KnownClientId.SPOManagementShell);
            }
            var body = $"resource={resourceUri}&client_id={clientId}&grant_type=password&username={HttpUtility.UrlEncode(username)}&password={HttpUtility.UrlEncode(password)}";
            var stringContent = new StringContent(body, System.Text.Encoding.UTF8, "application/x-www-form-urlencoded");

            var result = await client.PostAsync(tokenEndpoint, stringContent).ContinueWith<string>((response) =>
            {
                return response.Result.Content.ReadAsStringAsync().Result;
            });

            JObject jobject = JObject.Parse(result);

            // Ensure the resulting JSON could be parsed and that it doesn't contain an error. If incorrect credentials have been provided, this will not be the case and we return NULL to indicate not to have an access token.
            if (jobject == null || jobject["error"] != null)
            {
                var error = jobject["error"];

            }

            var token = jobject["access_token"].Value<string>();
            return token;
        }

        private void EnsureAzureADCredentialsToken(string resourceUri, string userPrincipalName, string userPassword, AzureEnvironment environment, KnownClientId clientId = KnownClientId.SPOManagementShell)
        {
            if (azureADCredentialsToken == null)
            {
                lock (tokenLock)
                {
                    if (azureADCredentialsToken == null)
                    {
                        var clientIdString = GetKnownClientId(clientId);
                        string accessToken = Task.Run(() => AcquireTokenAsync(resourceUri, userPrincipalName, userPassword, environment, clientIdString)).GetAwaiter().GetResult();

                        try
                        {
                            var token = new System.IdentityModel.Tokens.Jwt.JwtSecurityToken(accessToken);
                            Log.Debug(Constants.LOGGING_SOURCE, "Lease expiration date: {0}", token.ValidTo);
                            var lease = GetAccessTokenLease(token.ValidTo);
                            lease = TimeSpan.FromSeconds(lease.TotalSeconds - TimeSpan.FromMinutes(5).TotalSeconds > 0 ? lease.TotalSeconds - TimeSpan.FromMinutes(5).TotalSeconds : lease.TotalSeconds);

                            azureADCredentialsResetEvent = new AutoResetEvent(false);

                            AzureADCredentialsTokenWaitInfo wi = new AzureADCredentialsTokenWaitInfo();

                            wi.Handle = ThreadPool.RegisterWaitForSingleObject(azureADCredentialsResetEvent,
                                                                               new WaitOrTimerCallback(AzureADCredentialsTokenWaitProc),
                                                                               wi,
                                                                               (uint)lease.TotalMilliseconds,
                                                                               true);
                        }
                        catch (Exception ex)
                        {
                            Log.Warning(Constants.LOGGING_SOURCE, CoreResources.AuthenticationManger_ProblemDeterminingTokenLease, ex);
                            azureADCredentialsToken = null;
                        }

                        azureADCredentialsToken = accessToken;
                    }
                }
            }
        }

        internal class AzureADCredentialsTokenWaitInfo
        {
            public RegisteredWaitHandle Handle = null;
        }

        internal void AzureADCredentialsTokenWaitProc(object state, bool timedOut)
        {
            if (!timedOut)
            {
                AzureADCredentialsTokenWaitInfo wi = (AzureADCredentialsTokenWaitInfo)state;
                if (wi.Handle != null)
                {
                    wi.Handle.Unregister(null);
                }
            }
            else
            {
                azureADCredentialsToken = null;
            }
        }

        /// <summary>
        /// Returns a SharePoint ClientContext using Azure Active Directory authentication. This requires that you have a Azure AD Native Application registered. The user will be prompted for authentication.
        /// </summary>
        /// <param name="siteUrl">Site for which the ClientContext object will be instantiated</param>
        /// <param name="clientId">The Azure AD Native Application Client ID</param>
        /// <param name="redirectUrl">The Azure AD Native Application Redirect Uri as a string</param>
        /// <param name="tokenCache">Optional token cache. If not specified an in-memory token cache will be used</param>
        /// <param name="environment">SharePoint environment being used</param>
        /// <returns>Client context object</returns>
        public ClientContext GetAzureADNativeApplicationAuthenticatedContext(string siteUrl, string clientId, string redirectUrl, TokenCache tokenCache = null, AzureEnvironment environment = AzureEnvironment.Production)
        {
            return GetAzureADNativeApplicationAuthenticatedContext(siteUrl, clientId, new Uri(redirectUrl), tokenCache, environment);
        }

        /// <summary>
        /// Returns a SharePoint ClientContext using Azure Active Directory authentication. This requires that you have a Azure AD Native Application registered. The user will be prompted for authentication.
        /// </summary>
        /// <param name="siteUrl">Site for which the ClientContext object will be instantiated</param>
        /// <param name="clientId">The Azure AD Native Application Client ID</param>
        /// <param name="redirectUri">The Azure AD Native Application Redirect Uri</param>
        /// <param name="tokenCache">Optional token cache. If not specified an in-memory token cache will be used</param>
        /// <param name="environment">SharePoint environment being used</param>
        /// <returns>Client context object</returns>
        public ClientContext GetAzureADNativeApplicationAuthenticatedContext(string siteUrl, string clientId, Uri redirectUri, TokenCache tokenCache = null, AzureEnvironment environment = AzureEnvironment.Production)
        {
            var clientContext = new ClientContext(siteUrl);
            _contextUrl = siteUrl;
            _tokenCache = tokenCache;
            _clientId = clientId;
            _redirectUri = redirectUri;
            _commonAuthority = String.Format("{0}/common", GetAzureADLoginEndPoint(environment));

            clientContext.ExecutingWebRequest += clientContext_NativeApplicationExecutingWebRequest;

            return clientContext;
        }

        /// <summary>
        /// Returns a SharePoint ClientContext using Azure Active Directory authentication. This requires that you have a Azure AD Web Application registered. The user will not be prompted for authentication, the current user's authentication context will be used by leveraging ADAL.
        /// </summary>
        /// <param name="siteUrl">Site for which the ClientContext object will be instantiated</param>
        /// <param name="accessTokenGetter">The AccessToken getter method to use</param>
        /// <returns>Client context object</returns>
        public ClientContext GetAzureADWebApplicationAuthenticatedContext(String siteUrl, Func<String, String> accessTokenGetter)
        {
            var clientContext = new ClientContext(siteUrl)
            {
                DisableReturnValueCache = true
            };
            clientContext.ExecutingWebRequest += (sender, args) =>
            {
                Uri resourceUri = new Uri(siteUrl);
                resourceUri = new Uri(resourceUri.Scheme + "://" + resourceUri.Host + "/");

                String accessToken = accessTokenGetter(resourceUri.ToString());
                args.WebRequestExecutor.RequestHeaders["Authorization"] = "Bearer " + accessToken;
            };

            return clientContext;
        }

        /// <summary>
        /// Returns a SharePoint ClientContext using Azure Active Directory authentication. This requires that you have a Azure AD Web Application registered. The user will not be prompted for authentication, the current user's authentication context will be used by leveraging an explicit OAuth 2.0 Access Token value.
        /// </summary>
        /// <param name="siteUrl">Site for which the ClientContext object will be instantiated</param>
        /// <param name="accessToken">An explicit value for the AccessToken</param>
        /// <returns>Client context object</returns>
        public ClientContext GetAzureADAccessTokenAuthenticatedContext(String siteUrl, String accessToken)
        {
            var clientContext = new ClientContext(siteUrl)
            {
                DisableReturnValueCache = true
            };

            clientContext.ExecutingWebRequest += (sender, args) =>
            {
                args.WebRequestExecutor.RequestHeaders["Authorization"] = "Bearer " + accessToken;
            };

            return clientContext;
        }

        async void clientContext_NativeApplicationExecutingWebRequest(object sender, WebRequestEventArgs e)
        {
            var host = new Uri(_contextUrl);
            var ar = await AcquireNativeApplicationTokenAsync(_commonAuthority, host.Scheme + "://" + host.Host + "/");

            if (ar != null)
            {
                e.WebRequestExecutor.RequestHeaders["Authorization"] = "Bearer " + ar.AccessToken;
            }
        }

        private async Task<AuthenticationResult> AcquireNativeApplicationTokenAsync(string authContextUrl, string resourceId)
        {
            AuthenticationResult ar = null;

            await new SynchronizationContextRemover();

            try
            {
                if (_tokenCache != null)
                {
                    _authContext = new AuthenticationContext(authContextUrl, _tokenCache);
                }
                else
                {

                    _authContext = new AuthenticationContext(authContextUrl);
                }

                if (_authContext.TokenCache.ReadItems().Any())
                {
                    string cachedAuthority =
                        _authContext.TokenCache.ReadItems().First().Authority;

                    if (_tokenCache != null)
                    {
                        _authContext = new AuthenticationContext(cachedAuthority, _tokenCache);
                    }
                    else
                    {
                        _authContext = new AuthenticationContext(cachedAuthority);
                    }
                }
                ar = (await _authContext.AcquireTokenSilentAsync(resourceId, _clientId));
            }
            catch (Exception)
            {
                //not in cache; we'll get it with the full oauth flow
            }

            if (ar == null)
            {
                try
                {
                    ar = await _authContext.AcquireTokenAsync(resourceId, _clientId, _redirectUri, new PlatformParameters());
                }
                catch (Exception acquireEx)
                {
                    Log.Error(Constants.LOGGING_SOURCE, acquireEx.ToDetailedString());
                }
            }

            return ar;
        }

        /// <summary>
        /// Returns a SharePoint ClientContext using Azure Active Directory App Only Authentication. This requires that you have a certificated created, and updated the key credentials key in the application manifest in the azure AD accordingly.
        /// </summary>
        /// <param name="siteUrl">Site for which the ClientContext object will be instantiated</param>
        /// <param name="clientId">The Azure AD Application Client ID</param>
        /// <param name="tenant">The Azure AD Tenant, e.g. mycompany.onmicrosoft.com</param>
        /// <param name="storeName">The name of the store for the certificate</param>
        /// <param name="storeLocation">The location of the store for the certificate</param>
        /// <param name="thumbPrint">The thumbprint of the certificate to locate in the store</param>
        /// <param name="environment">SharePoint environment being used</param>
        /// <returns>ClientContext being used</returns>
        public ClientContext GetAzureADAppOnlyAuthenticatedContext(string siteUrl, string clientId, string tenant, StoreName storeName, StoreLocation storeLocation, string thumbPrint, AzureEnvironment environment = AzureEnvironment.Production)
        {
            var cert = Utilities.X509CertificateUtility.LoadCertificate(storeName, storeLocation, thumbPrint);

            return GetAzureADAppOnlyAuthenticatedContext(siteUrl, clientId, tenant, cert, environment);
        }


        /// <summary>
        /// Returns a SharePoint ClientContext using Azure Active Directory App Only Authentication. This requires that you have a certificated created, and updated the key credentials key in the application manifest in the azure AD accordingly.
        /// </summary>
        /// <param name="siteUrl">Site for which the ClientContext object will be instantiated</param>
        /// <param name="clientId">The Azure AD Application Client ID</param>
        /// <param name="tenant">The Azure AD Tenant, e.g. mycompany.onmicrosoft.com</param>
        /// <param name="certificatePath">The path to the certificate (*.pfx) file on the file system</param>
        /// <param name="certificatePassword">Password to the certificate</param>
        /// <param name="environment">SharePoint environment being used</param>
        /// <returns>Client context object</returns>
        public ClientContext GetAzureADAppOnlyAuthenticatedContext(string siteUrl, string clientId, string tenant, string certificatePath, string certificatePassword, AzureEnvironment environment = AzureEnvironment.Production)
        {
            var certPassword = Utilities.EncryptionUtility.ToSecureString(certificatePassword);

            return GetAzureADAppOnlyAuthenticatedContext(siteUrl, clientId, tenant, certificatePath, certPassword, environment);
        }


        /// <summary>
        /// Returns a SharePoint ClientContext using Azure Active Directory App Only Authentication. This requires that you have a certificated created, and updated the key credentials key in the application manifest in the azure AD accordingly.
        /// </summary>
        /// <param name="siteUrl">Site for which the ClientContext object will be instantiated</param>
        /// <param name="clientId">The Azure AD Application Client ID</param>
        /// <param name="tenant">The Azure AD Tenant, e.g. mycompany.onmicrosoft.com</param>
        /// <param name="certificatePath">The path to the certificate (*.pfx) file on the file system</param>
        /// <param name="certificatePassword">Password to the certificate</param>
        /// <param name="environment">SharePoint environment being used</param>
        /// <returns>Client context object</returns>
        public ClientContext GetAzureADAppOnlyAuthenticatedContext(string siteUrl, string clientId, string tenant, string certificatePath, SecureString certificatePassword, AzureEnvironment environment = AzureEnvironment.Production)
        {
            var certfile = System.IO.File.OpenRead(certificatePath);
            var certificateBytes = new byte[certfile.Length];
            certfile.Read(certificateBytes, 0, (int)certfile.Length);
            var cert = new X509Certificate2(
                certificateBytes,
                certificatePassword,
                X509KeyStorageFlags.Exportable |
                X509KeyStorageFlags.MachineKeySet |
                X509KeyStorageFlags.PersistKeySet);

            return GetAzureADAppOnlyAuthenticatedContext(siteUrl, clientId, tenant, cert, environment);
        }

        /// <summary>
        /// Returns a SharePoint ClientContext using Azure Active Directory App Only Authentication. This requires that you have a certificated created, and updated the key credentials key in the application manifest in the azure AD accordingly.
        /// </summary>
        /// <param name="siteUrl">Site for which the ClientContext object will be instantiated</param>
        /// <param name="clientId">The Azure AD Application Client ID</param>
        /// <param name="tenant">The Azure AD Tenant, e.g. mycompany.onmicrosoft.com</param>
        /// <param name="certificate">Certificate used to authenticate</param>
        /// <param name="environment">SharePoint environment being used</param>
        /// <returns></returns>
        public ClientContext GetAzureADAppOnlyAuthenticatedContext(string siteUrl, string clientId, string tenant, X509Certificate2 certificate, AzureEnvironment environment = AzureEnvironment.Production)
        {
            LoggerCallbackHandler.UseDefaultLogging = false;


            var clientContext = new ClientContext(siteUrl)
            {
                DisableReturnValueCache = true
            };

            string authority = string.Format(CultureInfo.InvariantCulture, "{0}/{1}/", GetAzureADLoginEndPoint(environment), tenant);

            var authContext = new AuthenticationContext(authority);

            var clientAssertionCertificate = new ClientAssertionCertificate(clientId, certificate);

            var host = new Uri(siteUrl);

            clientContext.ExecutingWebRequest += (sender, args) =>
            {
                var ar = Task.Run(() => authContext
                    .AcquireTokenAsync(host.Scheme + "://" + host.Host + "/", clientAssertionCertificate))
                    .GetAwaiter().GetResult();
                args.WebRequestExecutor.RequestHeaders["Authorization"] = "Bearer " + ar.AccessToken;
            };

            ClientContextSettings clientContextSettings = new ClientContextSettings()
            {
                Type = ClientContextType.AzureADCertificate,
                SiteUrl = siteUrl,
                AuthenticationManager = this,
                ClientId = clientId,
                Tenant = tenant,
                Certificate = certificate,
                Environment = environment
            };

            clientContext.AddContextSettings(clientContextSettings);

            return clientContext;
        }

        /// <summary>
        /// Returns a SharePoint ClientContext using Azure Active Directory App Only Authentication. This requires that you have a certificated created, and updated the key credentials key in the application manifest in the azure AD accordingly.
        /// </summary>
        /// <param name="siteUrl">Site for which the ClientContext object will be instantiated</param>
        /// <param name="clientId">The Azure AD Application Client ID</param>
        /// <param name="tenant">The Azure AD Tenant, e.g. mycompany.onmicrosoft.com</param>
        /// <param name="clientAssertionCertificate">IClientAssertionCertificate used to authenticate</param>
        /// <param name="environment">SharePoint environment being used</param>
        /// <returns></returns>
        public ClientContext GetAzureADAppOnlyAuthenticatedContext(string siteUrl, string clientId, string tenant, IClientAssertionCertificate clientAssertionCertificate, AzureEnvironment environment = AzureEnvironment.Production)
        {
            var clientContext = new ClientContext(siteUrl)
            {
                DisableReturnValueCache = true
            };

            string authority = string.Format(CultureInfo.InvariantCulture, "{0}/{1}/", GetAzureADLoginEndPoint(environment), tenant);

            var authContext = new AuthenticationContext(authority);

            //var clientAssertionCertificate = new ClientAssertionCertificate(clientId, certificate);

            var host = new Uri(siteUrl);

            clientContext.ExecutingWebRequest += (sender, args) =>
            {
                var ar = Task.Run(() => authContext
                    .AcquireTokenAsync(host.Scheme + "://" + host.Host + "/", clientAssertionCertificate))
                    .GetAwaiter().GetResult();
                args.WebRequestExecutor.RequestHeaders["Authorization"] = "Bearer " + ar.AccessToken;
            };

            ClientContextSettings clientContextSettings = new ClientContextSettings()
            {
                Type = ClientContextType.AzureADCertificate,
                SiteUrl = siteUrl,
                AuthenticationManager = this,
                ClientId = clientId,
                Tenant = tenant,
                ClientAssertionCertificate = clientAssertionCertificate,
                Environment = environment
            };

            clientContext.AddContextSettings(clientContextSettings);

            return clientContext;
        }

        /// <summary>
        /// Get's the Azure AD login end point for the given environment
        /// </summary>
        /// <param name="environment">Environment to get the login information for</param>
        /// <returns>Azure AD login endpoint</returns>
        public string GetAzureADLoginEndPoint(AzureEnvironment environment)
        {
            switch (environment)
            {
                case AzureEnvironment.Production:
                    {
                        return "https://login.microsoftonline.com";
                    }
                case AzureEnvironment.Germany:
                    {
                        return "https://login.microsoftonline.de";
                    }
                case AzureEnvironment.China:
                    {
                        return "https://login.chinacloudapi.cn";
                    }
                case AzureEnvironment.USGovernment:
                    {
                        return "https://login.microsoftonline.us";
                    }
                case AzureEnvironment.PPE:
                    {
                        return "https://login.windows-ppe.net";
                    }
                default:
                    {
                        return "https://login.microsoftonline.com";
                    }
            }
        }

        public static string GetSharePointDomainSuffix(AzureEnvironment environment)
        {
            if (environment == AzureEnvironment.Production)
            {
                return "com";
            }
            else if (environment == AzureEnvironment.USGovernment)
            {
                return "us";
            }
            else if (environment == AzureEnvironment.Germany)
            {
                return "de";
            }
            else if (environment == AzureEnvironment.China)
            {
                return "cn";
            }

            return "com";
        }

        protected virtual void Dispose(bool disposing)
        {
            if (!disposedValue)
            {
                if (disposing)
                {
                    if (appOnlyAccessTokenResetEvent != null)
                    {
                        appOnlyAccessTokenResetEvent.Set();
                        appOnlyAccessTokenResetEvent?.Dispose();
                    }

                    if (azureADCredentialsResetEvent != null)
                    {
                        azureADCredentialsResetEvent.Set();
                        azureADCredentialsResetEvent?.Dispose();
                    }
                }

                disposedValue = true;
            }
        }

        public void Dispose()
        {
            // Do not change this code. Put cleanup code in 'Dispose(bool disposing)' method
            Dispose(disposing: true);
            GC.SuppressFinalize(this);
        }
    }
}
