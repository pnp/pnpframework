using Microsoft.Identity.Client;
using Microsoft.SharePoint.Client;
using Newtonsoft.Json.Linq;
using PnP.Framework.Diagnostics;
using PnP.Framework.Utilities.Context;
using System;
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
        /// <summary>
        /// 
        /// </summary>
        Production = 0,
        /// <summary>
        /// 
        /// </summary>
        PPE = 1,
        /// <summary>
        /// 
        /// </summary>
        China = 2,
        /// <summary>
        /// 
        /// </summary>
        Germany = 3,
        /// <summary>
        /// 
        /// </summary>
        USGovernment = 4
    }

    /// <summary>
    /// This manager class can be used to obtain a SharePointContext object
    /// </summary>
    ///

    public enum KnownClientId
    {
        /// <summary>
        /// 
        /// </summary>
        PnPManagementShell,
        /// <summary>
        /// 
        /// </summary>
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
        private readonly object tokenLock = new object();
        private string _contextUrl;
        private string _commonAuthority = "https://login.windows.net/Common";
        private string _clientId;
        private Uri _redirectUri;
        private bool disposedValue;

        private IPublicClientApplication publicClientApplication;
        private IConfidentialClientApplication confidentialClientApplication;
        private string azureADEndPoint;
        private ClientContextType authenticationType;
        private string username;
        private X509Certificate2 certificate;
        private string clientSecret;
        private SecureString password;
        internal string RedirectUrl { get; set; }

        #region Construction
        public AuthenticationManager()
        {
            // Set the TLS preference. Needed on some server os's to work when Office 365 removes support for TLS 1.0
            ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12 | SecurityProtocolType.Tls11 | SecurityProtocolType.Tls;
        }

        /// <summary>
        /// /// Creates a new instance of the Authentication Manager to acquire authenticated ClientContexts. It uses the PnP Management Shell multi-tenant Azure AD application ID to authenticate.
        /// </summary>
        /// <param name="username"></param>
        /// <param name="password"></param>
        /// <param name="azureEnvironment"></param>
        public AuthenticationManager(string username, SecureString password, AzureEnvironment azureEnvironment = AzureEnvironment.Production) : this(GetKnownClientId(KnownClientId.PnPManagementShell), username, password, "https://login.microsoftonline.com/common/oauth2/nativeclient", azureEnvironment)
        {
        }

        /// <summary>
        /// Creates a new instance of the Authentication Manager to acquire authenticated ClientContexts.
        /// </summary>
        /// <param name="clientId"></param>
        /// <param name="username"></param>
        /// <param name="password"></param>
        /// <param name="redirectUrl"></param>
        /// <param name="azureEnvironment"></param>
        public AuthenticationManager(string clientId, string username, SecureString password, string redirectUrl = null, AzureEnvironment azureEnvironment = AzureEnvironment.Production)
        {
            azureADEndPoint = GetAzureADLoginEndPoint(azureEnvironment);

            var builder = PublicClientApplicationBuilder.Create(clientId).WithAuthority($"{azureADEndPoint}/organizations/");
            if (!string.IsNullOrEmpty(redirectUrl))
            {
                builder = builder.WithRedirectUri(redirectUrl);
            }
            this.username = username;
            this.password = password;
            publicClientApplication = builder.Build();
            authenticationType = ClientContextType.AzureADCredentials;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="clientId"></param>
        /// <param name="redirectUrl"></param>
        /// <param name="azureEnvironment"></param>
        public AuthenticationManager(string clientId, string redirectUrl = null, AzureEnvironment azureEnvironment = AzureEnvironment.Production)
        {
            azureADEndPoint = GetAzureADLoginEndPoint(azureEnvironment);
            var builder = PublicClientApplicationBuilder.Create(clientId).WithAuthority($"{azureADEndPoint}/organizations/");
            if (!string.IsNullOrEmpty(redirectUrl))
            {
                builder = builder.WithRedirectUri(redirectUrl);
            }
            publicClientApplication = builder.Build();
            authenticationType = ClientContextType.AzureADInteractive;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="clientId"></param>
        /// <param name="certificate"></param>
        /// <param name="redirectUrl"></param>
        /// <param name="azureEnvironment"></param>
        public AuthenticationManager(string clientId, X509Certificate2 certificate, string redirectUrl = null, AzureEnvironment azureEnvironment = AzureEnvironment.Production)
        {
            azureADEndPoint = GetAzureADLoginEndPoint(azureEnvironment);
            var builder = ConfidentialClientApplicationBuilder.Create(clientId).WithCertificate(certificate).WithAuthority($"{azureADEndPoint}/organizations/");
            if (!string.IsNullOrEmpty(redirectUrl))
            {
                builder = builder.WithRedirectUri(redirectUrl);
            }
            confidentialClientApplication = builder.Build();
            this.certificate = certificate;
            authenticationType = ClientContextType.AzureADCredentials;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="clientId"></param>
        /// <param name="certificatePath"></param>
        /// <param name="certificatePassword"></param>
        /// <param name="redirectUrl"></param>
        /// <param name="azureEnvironment"></param>
        public AuthenticationManager(string clientId, string certificatePath, string certificatePassword, string redirectUrl = null, AzureEnvironment azureEnvironment = AzureEnvironment.Production)
        {
            azureADEndPoint = GetAzureADLoginEndPoint(azureEnvironment);
            var builder = ConfidentialClientApplicationBuilder.Create(clientId).WithAuthority($"{azureADEndPoint}/organizations/");
            if (!string.IsNullOrEmpty(redirectUrl))
            {
                builder = builder.WithRedirectUri(redirectUrl);
            }
            confidentialClientApplication = builder.Build();

            var certfile = System.IO.File.OpenRead(certificatePath);
            var certificateBytes = new byte[certfile.Length];
            certfile.Read(certificateBytes, 0, (int)certfile.Length);
            var certificate = new X509Certificate2(
                certificateBytes,
                certificatePassword,
                X509KeyStorageFlags.Exportable |
                X509KeyStorageFlags.MachineKeySet |
                X509KeyStorageFlags.PersistKeySet);
            this.certificate = certificate;
            authenticationType = ClientContextType.AzureADCertificate;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="clientId"></param>
        /// <param name="storeName"></param>
        /// <param name="storeLocation"></param>
        /// <param name="thumbPrint"></param>
        /// <param name="redirectUrl"></param>
        /// <param name="azureEnvironment"></param>
        public AuthenticationManager(string clientId, StoreName storeName, StoreLocation storeLocation, string thumbPrint, string redirectUrl = null, AzureEnvironment azureEnvironment = AzureEnvironment.Production)
        {
            azureADEndPoint = GetAzureADLoginEndPoint(azureEnvironment);
            var builder = ConfidentialClientApplicationBuilder.Create(clientId).WithAuthority($"{azureADEndPoint}/organizations/");
            if (!string.IsNullOrEmpty(redirectUrl))
            {
                builder = builder.WithRedirectUri(redirectUrl);
            }
            confidentialClientApplication = builder.Build();

            this.certificate = Utilities.X509CertificateUtility.LoadCertificate(storeName, storeLocation, thumbPrint); ;
            authenticationType = ClientContextType.AzureADCertificate;
        }
        #endregion

        /// <summary>
        /// 
        /// </summary>
        /// <param name="siteUrl"></param>
        /// <returns></returns>
        public ClientContext GetContext(string siteUrl)
        {
            return GetContextAsync(siteUrl).GetAwaiter().GetResult();
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="siteUrl"></param>
        /// <returns></returns>
        public async Task<ClientContext> GetContextAsync(string siteUrl)
        {
            var uri = new Uri(siteUrl);

            var scopes = new[] { $"{uri.Scheme}://{uri.Authority}/.default" };

            AuthenticationResult authResult = null;
            switch (authenticationType)
            {
                case ClientContextType.AzureADCredentials:
                    {
                        var accounts = await publicClientApplication.GetAccountsAsync();
                        try
                        {
                            authResult = await publicClientApplication.AcquireTokenSilent(scopes, accounts.First()).ExecuteAsync();
                        }
                        catch
                        {
                            authResult = await publicClientApplication.AcquireTokenByUsernamePassword(scopes, username, password).ExecuteAsync();
                        }
                        if (authResult.AccessToken != null)
                        {
                            return BuildClientContext(publicClientApplication, siteUrl, scopes, authenticationType);
                        }
                        break;
                    }
                case ClientContextType.AzureADInteractive:
                    {
                        var accounts = await publicClientApplication.GetAccountsAsync();

                        try
                        {
                            authResult = await publicClientApplication.AcquireTokenSilent(scopes, accounts.First()).ExecuteAsync();
                        }
                        catch
                        {
                            authResult = await publicClientApplication.AcquireTokenInteractive(scopes).ExecuteAsync();
                        }
                        if (authResult.AccessToken != null)
                        {
                            return BuildClientContext(publicClientApplication, siteUrl, scopes, authenticationType);
                        }
                        break;
                    }
                case ClientContextType.AzureADCertificate:
                    {
                        var accounts = await confidentialClientApplication.GetAccountsAsync();

                        try
                        {
                            authResult = await confidentialClientApplication.AcquireTokenSilent(scopes, accounts.First()).ExecuteAsync();
                        }
                        catch
                        {
                            authResult = await confidentialClientApplication.AcquireTokenForClient(scopes).ExecuteAsync();
                        }
                        if (authResult.AccessToken != null)
                        {
                            return BuildClientContext(confidentialClientApplication, siteUrl, scopes, authenticationType);
                        }
                        break;
                    }
            }
            return null;
        }


        private ClientContext BuildClientContext(IClientApplicationBase application, string siteUrl, string[] scopes, ClientContextType contextType)
        {
            var clientContext = new ClientContext(siteUrl)
            {
                DisableReturnValueCache = true
            };

            clientContext.ExecutingWebRequest += (sender, args) =>
            {
                var accounts = application.GetAccountsAsync().GetAwaiter().GetResult();
                var ar = application.AcquireTokenSilent(scopes, accounts.First()).ExecuteAsync().GetAwaiter().GetResult();
                args.WebRequestExecutor.RequestHeaders["Authorization"] = "Bearer " + ar.AccessToken;
            };

            ClientContextSettings clientContextSettings = new ClientContextSettings()
            {
                Type = contextType,
                SiteUrl = siteUrl,
                AuthenticationManager = this,
            };

            clientContext.AddContextSettings(clientContextSettings);

            return clientContext;
        }

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
        public ClientContext GetACSAppOnlyContext(string siteUrl, string appId, string appSecret)
        {
            return GetACSAppOnlyContext(siteUrl, Utilities.TokenHelper.GetRealmFromTargetUrl(new Uri(siteUrl)), appId, appSecret);
        }

        /// <summary>
        /// Returns an app only ClientContext object
        /// </summary>
        /// <param name="siteUrl">Site for which the ClientContext object will be instantiated</param>
        /// <param name="appId">Application ID which is requesting the ClientContext object</param>
        /// <param name="appSecret">Application secret of the Application which is requesting the ClientContext object</param>
        /// <param name="environment">SharePoint environment being used</param>
        /// <returns>ClientContext to be used by CSOM code</returns>
        public ClientContext GetACSAppOnlyContext(string siteUrl, string appId, string appSecret, AzureEnvironment environment = AzureEnvironment.Production)
        {
            return GetACSAppOnlyContext(siteUrl, Utilities.TokenHelper.GetRealmFromTargetUrl(new Uri(siteUrl)), appId, appSecret, GetACSEndPoint(environment), GetACSEndPointPrefix(environment));
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
        public ClientContext GetACSAppOnlyContext(string siteUrl, string realm, string appId, string appSecret, string acsHostUrl = "accesscontrol.windows.net", string globalEndPointPrefix = "accounts")
        {
            ACSEnsureToken(siteUrl, realm, appId, appSecret, acsHostUrl, globalEndPointPrefix);
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
        public string GetACSEndPoint(AzureEnvironment environment)
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
        public string GetACSEndPointPrefix(AzureEnvironment environment)
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
        private void ACSEnsureToken(string siteUrl, string realm, string appId, string appSecret, string acsHostUrl, string globalEndPointPrefix)
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

        ///// <summary>
        ///// Returns a SharePoint ClientContext using Azure Active Directory credential authentication. This depends on the SPO Management Shell app or the PnP Management Shell app being registered in your Azure AD.
        ///// </summary>
        ///// <param name="siteUrl">Site for which the ClientContext object will be instantiated</param>
        ///// <param name="userPrincipalName">The user id</param>
        ///// <param name="userPassword">The user's password as a secure string</param>
        ///// <param name="environment">SharePoint environment being used</param>
        ///// <param name="clientId">Enum value pointing to one of the known client ids</param></parm>
        ///// <returns>Client context object</returns>
        //public ClientContext GetCredentialsContext(string siteUrl, string userPrincipalName, SecureString userPassword, AzureEnvironment environment = AzureEnvironment.Production, KnownClientId clientId = KnownClientId.SPOManagementShell)
        //{
        //    string password = new System.Net.NetworkCredential(string.Empty, userPassword).Password;
        //    return GetCredentialsContext(siteUrl, userPrincipalName, password, environment, clientId);
        //}

        ///// <summary>
        ///// Returns a SharePoint ClientContext using Azure Active Directory credential authentication. This depends on the SPO Management Shell app or the PnP Management Shell app being registered in your Azure AD.
        ///// </summary>
        ///// <param name="siteUrl">Site for which the ClientContext object will be instantiated</param>
        ///// <param name="userPrincipalName">The user id</param>
        ///// <param name="userPassword">The user's password as a string</param>
        ///// <param name="environment">SharePoint environment being used</param>
        ///// <returns>Client context object</returns>
        //public ClientContext GetCredentialsContext(string siteUrl, string userPrincipalName, string userPassword, AzureEnvironment environment = AzureEnvironment.Production, KnownClientId clientId = KnownClientId.SPOManagementShell)
        //{
        //    Log.Info(Constants.LOGGING_SOURCE, CoreResources.AuthenticationManager_GetContext, siteUrl);
        //    Log.Debug(Constants.LOGGING_SOURCE, CoreResources.AuthenticationManager_TenantUser, userPrincipalName);

        //    var spUri = new Uri(siteUrl);
        //    string resourceUri = spUri.Scheme + "://" + spUri.Authority;

        //    var clientContext = new ClientContext(siteUrl)
        //    {
        //        DisableReturnValueCache = true
        //    };
        //    clientContext.ExecutingWebRequest += (sender, args) =>
        //    {
        //        EnsureAzureADCredentialsToken(resourceUri, userPrincipalName, userPassword, environment, clientId);
        //        args.WebRequestExecutor.RequestHeaders["Authorization"] = "Bearer " + azureADCredentialsToken;
        //    };

        //    ClientContextSettings clientContextSettings = new ClientContextSettings()
        //    {
        //        Type = ClientContextType.AzureADCredentials,
        //        SiteUrl = siteUrl,
        //        AuthenticationManager = this,
        //        UserName = userPrincipalName,
        //        Password = userPassword
        //    };

        //    clientContext.AddContextSettings(clientContextSettings);

        //    return clientContext;
        //}

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

        ///// <summary>
        ///// Returns a SharePoint ClientContext using Azure Active Directory authentication. This requires that you have a Azure AD Native Application registered. The user will be prompted for authentication.
        ///// </summary>
        ///// <param name="siteUrl">Site for which the ClientContext object will be instantiated</param>
        ///// <param name="clientId">The Azure AD Native Application Client ID</param>
        ///// <param name="redirectUrl">The Azure AD Native Application Redirect Uri as a string</param>
        ///// <param name="tokenCache">Optional token cache. If not specified an in-memory token cache will be used</param>
        ///// <param name="environment">SharePoint environment being used</param>
        ///// <returns>Client context object</returns>
        //public ClientContext GetInteractiveContext(string siteUrl, string clientId, string redirectUrl, AzureEnvironment environment = AzureEnvironment.Production)
        //{
        //    return GetInteractiveContext(siteUrl, clientId, new Uri(redirectUrl), environment);
        //}

        ///// <summary>
        ///// Returns a SharePoint ClientContext using Azure Active Directory authentication. This requires that you have a Azure AD Native Application registered. The user will be prompted for authentication.
        ///// </summary>
        ///// <param name="siteUrl">Site for which the ClientContext object will be instantiated</param>
        ///// <param name="clientId">The Azure AD Native Application Client ID</param>
        ///// <param name="redirectUri">The Azure AD Native Application Redirect Uri</param>
        ///// <param name="environment">SharePoint environment being used</param>
        ///// <returns>Client context object</returns>
        //public ClientContext GetInteractiveContext(string siteUrl, string clientId, Uri redirectUri, AzureEnvironment environment = AzureEnvironment.Production)
        //{
        //    var clientContext = new ClientContext(siteUrl);
        //    _contextUrl = siteUrl;
        //    _clientId = clientId;
        //    _redirectUri = redirectUri;
        //    _commonAuthority = String.Format("{0}/common", GetAzureADLoginEndPoint(environment));

        //    clientContext.ExecutingWebRequest += clientContext_ApplicationTokenInteractiveExecutingWebRequest;

        //    return clientContext;
        //}

        /// <summary>
        /// Returns a SharePoint ClientContext using Azure Active Directory authentication. This requires that you have a Azure AD Web Application registered. The user will not be prompted for authentication, the current user's authentication context will be used by leveraging ADAL.
        /// </summary>
        /// <param name="siteUrl">Site for which the ClientContext object will be instantiated</param>
        /// <param name="accessTokenGetter">The AccessToken getter method to use</param>
        /// <returns>Client context object</returns>
        public ClientContext GetAzureADWebApplicationAuthenticatedContext(string siteUrl, Func<string, string> accessTokenGetter)
        {
            var clientContext = new ClientContext(siteUrl)
            {
                DisableReturnValueCache = true
            };
            clientContext.ExecutingWebRequest += (sender, args) =>
            {
                Uri resourceUri = new Uri(siteUrl);
                resourceUri = new Uri(resourceUri.Scheme + "://" + resourceUri.Host + "/");

                string accessToken = accessTokenGetter(resourceUri.ToString());
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
        public ClientContext GetAccessTokenAuthenticatedContext(string siteUrl, string accessToken)
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
