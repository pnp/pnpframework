using Microsoft.Identity.Client;
using Microsoft.Identity.Client.Extensibility;
using Microsoft.SharePoint.Client;
using PnP.Core.Services;
using PnP.Framework.Utilities;
using PnP.Framework.Utilities.Context;
using System;
using System.Configuration;
using System.Linq;
using System.Net;
using System.Runtime.InteropServices;
using System.Security;
using System.Security.Cryptography.X509Certificates;
using System.Threading;
using System.Threading.Tasks;

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
        USGovernment = 4,
        /// <summary>
        /// 
        /// </summary>
        USGovernmentHigh = 5,
        /// <summary>
        /// 
        /// </summary>
        USGovernmentDoD = 6,

        /// <summary>
        /// Custom cloud configuration, specify the endpoints manually
        /// </summary>
        Custom = 100
    }


    /// <summary>
    /// A Known Client Ids to use for authentication
    /// </summary>
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

    /// <summary>
    /// This manager class can be used to obtain a SharePoint Client Context object
    /// </summary>
    public class AuthenticationManager : IDisposable
    {
        /// <summary>
        /// The client id of the Microsoft SharePoint Online Management Shell application
        /// </summary>
        public const string CLIENTID_SPOMANAGEMENTSHELL = "9bc3ab49-b65d-410a-85ad-de819febfddc";
        /// <summary>
        /// The client id of the Microsoft 365 Patters and Practices Management Shell application
        /// </summary>
        public const string CLIENTID_PNPMANAGEMENTSHELL = "31359c7f-bd7e-475c-86db-fdb8c937548e";

        private readonly IPublicClientApplication publicClientApplication;
        private readonly IConfidentialClientApplication confidentialClientApplication;

        // Azure environment setup
        private AzureEnvironment azureEnvironment;
        // When azureEnvironment = Custom then use these strings to keep track of the respective URLs to use 
        private string microsoftGraphEndPoint;
        private string azureADLoginEndPoint;

        private readonly ClientContextType authenticationType;
        private readonly string username;
        private readonly SecureString password;
        private readonly UserAssertion assertion;
        private readonly Func<DeviceCodeResult, Task> deviceCodeCallback;
        private readonly ICustomWebUi customWebUi;
        private readonly ACSTokenGenerator acsTokenGenerator;
        private IMsalHttpClientFactory httpClientFactory;
        private readonly SecureString accessToken;
        private readonly IAuthenticationProvider authenticationProvider;
        private readonly PnPContext pnpContext;

        public CookieContainer CookieContainer { get; set; }

        private IMsalHttpClientFactory HttpClientFactory
        {
            get
            {
                if (httpClientFactory == null)
                {
                    httpClientFactory = new Http.MsalHttpClientFactory();
                }
                return httpClientFactory;
            }
        }

        #region Creation

        public static AuthenticationManager CreateWithAccessToken(SecureString accessToken)
        {
            return new AuthenticationManager(accessToken);
        }

        /// <summary>
        /// Creates a new instance of the Authentication Manager to acquire authenticated ClientContexts through device code authentication
        /// </summary>
        /// <param name="clientId">The client id of the Azure AD application to use for authentication</param>
        /// <param name="deviceCodeCallback">The callback that will be called with device code information.</param>
        /// <param name="azureEnvironment">The azure environment to use. Defaults to AzureEnvironment.Production</param>
        /// <param name="tokenCacheCallback">If present, after setting up the base flow for authentication this callback will be called register a custom tokencache. See https://aka.ms/msal-net-token-cache-serialization.</param>
        public static AuthenticationManager CreateWithDeviceLogin(string clientId, Func<DeviceCodeResult, Task> deviceCodeCallback, AzureEnvironment azureEnvironment = AzureEnvironment.Production, Action<ITokenCache> tokenCacheCallback = null)
        {
            return new AuthenticationManager(clientId, null, deviceCodeCallback, azureEnvironment, tokenCacheCallback);
        }

        /// <summary>
        /// Creates a new instance of the Authentication Manager to acquire authenticated ClientContexts through device code authentication
        /// </summary>
        /// <param name="clientId">The client id of the Azure AD application to use for authentication</param>
        /// <param name="tenantId">Optional tenant id or tenant url</param>
        /// <param name="deviceCodeCallback">The callback that will be called with device code information.</param>
        /// <param name="azureEnvironment">The azure environment to use. Defaults to AzureEnvironment.Production</param>
        /// <param name="tokenCacheCallback">If present, after setting up the base flow for authentication this callback will be called register a custom tokencache. See https://aka.ms/msal-net-token-cache-serialization.</param>
        public static AuthenticationManager CreateWithDeviceLogin(string clientId, string tenantId, Func<DeviceCodeResult, Task> deviceCodeCallback, AzureEnvironment azureEnvironment = AzureEnvironment.Production, Action<ITokenCache> tokenCacheCallback = null)
        {
            return new AuthenticationManager(clientId, tenantId, deviceCodeCallback, azureEnvironment, tokenCacheCallback);
        }

        /// <summary>
        /// Creates a new instance of the Authentication Manager to acquire access tokens and client contexts using the Azure AD Interactive flow.
        /// </summary>
        /// <param name="clientId">The client id of the Azure AD application to use for authentication</param>
        /// <param name="openBrowserCallback">This callback will be called providing the URL and port to open during the authentication flow</param>
        /// <param name="tenantId">Optional tenant id or tenant url</param>
        /// <param name="successMessageHtml">Allows you to override the success message. Notice that a success header message will be added.</param>
        /// <param name="failureMessageHtml">llows you to override the failure message. Notice that a failed header message will be added and the error message will be appended.</param>
        /// <param name="azureEnvironment">The azure environment to use. Defaults to AzureEnvironment.Production</param>
        /// <param name="tokenCacheCallback">If present, after setting up the base flow for authentication this callback will be called to register a custom tokencache. See https://aka.ms/msal-net-token-cache-serialization.</param>
        public static AuthenticationManager CreateWithInteractiveLogin(string clientId, Action<string, int> openBrowserCallback, string tenantId = null, string successMessageHtml = null, string failureMessageHtml = null, AzureEnvironment azureEnvironment = AzureEnvironment.Production, Action<ITokenCache> tokenCacheCallback = null)
        {
            return new AuthenticationManager(clientId, Utilities.OAuth.DefaultBrowserUi.FindFreeLocalhostRedirectUri(), tenantId, azureEnvironment, tokenCacheCallback, new Utilities.OAuth.DefaultBrowserUi(openBrowserCallback, successMessageHtml, failureMessageHtml));
        }

        /// <summary>
        /// Creates a new instance of the Authentication Manager to acquire access tokens and client contexts using the Azure AD Interactive flow.
        /// </summary>
        /// <param name="clientId">The client id of the Azure AD application to use for authentication</param>
        /// <param name="redirectUrl">Optional redirect URL to use for authentication as set up in the Azure AD Application</param>
        /// <param name="tenantId">Optional tenant id or tenant url</param>
        /// <param name="azureEnvironment">The azure environment to use. Defaults to AzureEnvironment.Production</param>
        /// <param name="tokenCacheCallback">If present, after setting up the base flow for authentication this callback will be called register a custom tokencache. See https://aka.ms/msal-net-token-cache-serialization.</param>
        /// <param name="customWebUi">Optional ICustomWebUi object to fully customize the feedback behavior</param>
        public static AuthenticationManager CreateWithInteractiveLogin(string clientId, string redirectUrl = null, string tenantId = null, AzureEnvironment azureEnvironment = AzureEnvironment.Production, Action<ITokenCache> tokenCacheCallback = null, ICustomWebUi customWebUi = null)
        {
            return new AuthenticationManager(clientId, redirectUrl ?? Utilities.OAuth.DefaultBrowserUi.FindFreeLocalhostRedirectUri(), tenantId, azureEnvironment, tokenCacheCallback, customWebUi);
        }

        /// <summary>
        /// Creates a new instance of the Authentication Manager to acquire authenticated ClientContexts. It uses the PnP Management Shell multi-tenant Azure AD application ID to authenticate. By default tokens will be cached in memory.
        /// </summary>
        /// <param name="username">The username to use for authentication</param>
        /// <param name="password">The password to use for authentication</param>
        /// <param name="azureEnvironment">The azure environment to use. Defaults to AzureEnvironment.Production</param>
        /// <param name="tokenCacheCallback">If present, after setting up the base flow for authentication this callback will be called register a custom tokencache. See https://aka.ms/msal-net-token-cache-serialization.</param>
        public static AuthenticationManager CreateWithCredentials(string username, SecureString password, AzureEnvironment azureEnvironment = AzureEnvironment.Production, Action<ITokenCache> tokenCacheCallback = null)
        {
            return new AuthenticationManager(username, password, azureEnvironment, tokenCacheCallback);
        }

        /// <summary>
        /// Creates a new instance of the Authentication Manager to acquire authenticated ClientContexts.
        /// </summary>
        /// <param name="clientId">The client id of the Azure AD application to use for authentication</param>
        /// <param name="username">The username to use for authentication</param>
        /// <param name="password">The password to use for authentication</param>
        /// <param name="redirectUrl">Optional redirect URL to use for authentication as set up in the Azure AD Application</param>
        /// <param name="azureEnvironment">The azure environment to use. Defaults to AzureEnvironment.Production</param>
        /// <param name="tokenCacheCallback">If present, after setting up the base flow for authentication this callback will be called register a custom tokencache. See https://aka.ms/msal-net-token-cache-serialization.</param>
        public static AuthenticationManager CreateWithCredentials(string clientId, string username, SecureString password, string redirectUrl = null, AzureEnvironment azureEnvironment = AzureEnvironment.Production, Action<ITokenCache> tokenCacheCallback = null)
        {
            return new AuthenticationManager(clientId, username, password, redirectUrl, azureEnvironment, tokenCacheCallback);
        }

        /// <summary>
        /// Creates a new instance of the Authentication Manager to acquire authenticated ClientContexts.
        /// </summary>
        /// <param name="clientId">The client id of the Azure AD application to use for authentication</param>
        /// <param name="certificate">A valid certificate</param>
        /// <param name="tenantId">Tenant id or tenant url</param>
        /// <param name="redirectUrl">Optional redirect URL to use for authentication as set up in the Azure AD Application</param>
        /// <param name="azureEnvironment">The azure environment to use. Defaults to AzureEnvironment.Production</param>
        /// <param name="tokenCacheCallback">If present, after setting up the base flow for authentication this callback will be called register a custom tokencache. See https://aka.ms/msal-net-token-cache-serialization.</param>
        public static AuthenticationManager CreateWithCertificate(string clientId, X509Certificate2 certificate, string tenantId, string redirectUrl = null, AzureEnvironment azureEnvironment = AzureEnvironment.Production, Action<ITokenCache> tokenCacheCallback = null)
        {
            return new AuthenticationManager(clientId, certificate, tenantId, redirectUrl, azureEnvironment, tokenCacheCallback);
        }

        /// <summary>
        /// Creates a new instance of the Authentication Manager to acquire authenticated ClientContexts.
        /// </summary>
        /// <param name="clientId">The client id of the Azure AD application to use for authentication</param>
        /// <param name="certificatePath">A valid path to a certificate file</param>
        /// <param name="certificatePassword">The password for the certificate</param>
        /// <param name="tenantId">The tenant id (guid) or name (e.g. contoso.onmicrosoft.com) </param>
        /// <param name="redirectUrl">Optional redirect URL to use for authentication as set up in the Azure AD Application</param>
        /// <param name="azureEnvironment">The azure environment to use. Defaults to AzureEnvironment.Production</param>
        /// <param name="tokenCacheCallback">If present, after setting up the base flow for authentication this callback will be called register a custom tokencache. See https://aka.ms/msal-net-token-cache-serialization.</param>
        public static AuthenticationManager CreateWithCertificate(string clientId, string certificatePath, string certificatePassword, string tenantId, string redirectUrl = null, AzureEnvironment azureEnvironment = AzureEnvironment.Production, Action<ITokenCache> tokenCacheCallback = null)
        {
            return new AuthenticationManager(clientId, certificatePath, certificatePassword, tenantId, redirectUrl, azureEnvironment, tokenCacheCallback);
        }

        /// <summary>
        /// Creates a new instance of the Authentication Manager to acquire authenticated ClientContexts.
        /// </summary>
        /// <param name="clientId">The client id of the Azure AD application to use for authentication</param>
        /// <param name="storeName">The name of the certificate store to find the certificate in.</param>
        /// <param name="storeLocation">The location of the certificate store to find the certificate in.</param>
        /// <param name="thumbPrint">The thumbprint of the certificate to use.</param>
        /// <param name="tenantId">The tenant id (guid) or name (e.g. contoso.onmicrosoft.com) </param>
        /// <param name="redirectUrl">Optional redirect URL to use for authentication as set up in the Azure AD Application</param>
        /// <param name="azureEnvironment">The azure environment to use. Defaults to AzureEnvironment.Production</param>
        /// <param name="tokenCacheCallback">If present, after setting up the base flow for authentication this callback will be called register a custom tokencache. See https://aka.ms/msal-net-token-cache-serialization.</param>
        public static AuthenticationManager CreateWithCertificate(string clientId, StoreName storeName, StoreLocation storeLocation, string thumbPrint, string tenantId, string redirectUrl = null, AzureEnvironment azureEnvironment = AzureEnvironment.Production, Action<ITokenCache> tokenCacheCallback = null)
        {
            return new AuthenticationManager(clientId, storeName, storeLocation, thumbPrint, tenantId, redirectUrl, azureEnvironment, tokenCacheCallback);
        }

        /// <summary>
        /// Creates a new instance of the Authentication Manager to acquire authenticated ClientContext.
        /// </summary>
        /// <param name="clientId">The client id of the Azure AD application to use for authentication.</param>
        /// <param name="clientSecret">The client secret of the Azure AD application to use for authentication.</param>
        /// <param name="tenantId">Optional tenant id or tenant url</param>
        /// <param name="azureEnvironment">The azure environment to use. Defaults to AzureEnvironment.Production</param>
        /// <param name="userAssertion">The user assertion (token) of the user on whose behalf to acquire the context</param>
        /// <param name="tokenCacheCallback">If present, after setting up the base flow for authentication this callback will be called register a custom tokencache. See https://aka.ms/msal-net-token-cache-serialization.</param>
        public static AuthenticationManager CreateWithOnBehalfOf(string clientId, string clientSecret, UserAssertion userAssertion, string tenantId = null, AzureEnvironment azureEnvironment = AzureEnvironment.Production, Action<ITokenCache> tokenCacheCallback = null)
        {
            return new AuthenticationManager(clientId, clientSecret, userAssertion, tenantId, azureEnvironment, tokenCacheCallback);
        }

        /// <summary>
        /// Creates a new instance of the Authentication Manager to acquire an authenticated ClientContext.
        /// </summary>
        /// <param name="authenticationProvider">PnP Core SDK authentication provider that will deliver the access token</param>
        /// <returns></returns>
        public static AuthenticationManager CreateWithPnPCoreSdk(IAuthenticationProvider authenticationProvider)
        {
            return new AuthenticationManager(authenticationProvider);
        }

        /// <summary>
        /// Creates a new instance of the Authentication Manager to acquire an authenticated ClientContext.
        /// </summary>
        /// <param name="pnpContext">PnP Core SDK authentication provider that will deliver the access token</param>
        /// <returns></returns>
        public static AuthenticationManager CreateWithPnPCoreSdk(PnPContext pnpContext)
        {
            return new AuthenticationManager(pnpContext);
        }
        #endregion

        #region Construction
        /// <summary>
        /// Empty constructor, to be used if you want to execute ACS based authentication methods.
        /// </summary>
        public AuthenticationManager()
        {
            // Set the TLS preference. Needed on some server os's to work when Office 365 removes support for TLS 1.0
            ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12 | SecurityProtocolType.Tls11 | SecurityProtocolType.Tls;
        }

        private AuthenticationManager(ACSTokenGenerator oAuthAuthenticationProvider) : this()
        {
            this.acsTokenGenerator = oAuthAuthenticationProvider;
            authenticationType = ClientContextType.SharePointACSAppOnly;
        }


        public AuthenticationManager(SecureString accessToken)
        {
            this.accessToken = accessToken;
            authenticationType = ClientContextType.AccessToken;
        }
        /// <summary>
        /// Creates a new instance of the Authentication Manager to acquire authenticated ClientContexts. It uses the PnP Management Shell multi-tenant Azure AD application ID to authenticate. By default tokens will be cached in memory.
        /// </summary>
        /// <param name="username">The username to use for authentication</param>
        /// <param name="password">The password to use for authentication</param>
        /// <param name="azureEnvironment">The azure environment to use. Defaults to AzureEnvironment.Production</param>
        /// <param name="tokenCacheCallback">If present, after setting up the base flow for authentication this callback will be called register a custom tokencache. See https://aka.ms/msal-net-token-cache-serialization.</param>
        public AuthenticationManager(string username, SecureString password, AzureEnvironment azureEnvironment = AzureEnvironment.Production, Action<ITokenCache> tokenCacheCallback = null) : this(GetKnownClientId(KnownClientId.PnPManagementShell), username, password, $"{GetAzureADLoginEndPointStatic(azureEnvironment)}/common/oauth2/nativeclient", azureEnvironment, tokenCacheCallback)
        {
        }

        /// <summary>
        /// Creates a new instance of the Authentication Manager to acquire authenticated ClientContexts.
        /// </summary>
        /// <param name="clientId">The client id of the Azure AD application to use for authentication</param>
        /// <param name="username">The username to use for authentication</param>
        /// <param name="password">The password to use for authentication</param>
        /// <param name="redirectUrl">Optional redirect URL to use for authentication as set up in the Azure AD Application</param>
        /// <param name="azureEnvironment">The azure environment to use. Defaults to AzureEnvironment.Production</param>
        /// <param name="tokenCacheCallback">If present, after setting up the base flow for authentication this callback will be called register a custom tokencache. See https://aka.ms/msal-net-token-cache-serialization.</param>
        public AuthenticationManager(string clientId, string username, SecureString password, string redirectUrl = null, AzureEnvironment azureEnvironment = AzureEnvironment.Production, Action<ITokenCache> tokenCacheCallback = null) : this()
        {
            this.azureEnvironment = azureEnvironment;
            var azureADEndPoint = GetAzureADLoginEndPoint(azureEnvironment);

            var builder = PublicClientApplicationBuilder.Create(clientId).WithAuthority($"{azureADEndPoint}/organizations/").WithHttpClientFactory(HttpClientFactory);
            if (!string.IsNullOrEmpty(redirectUrl))
            {
                builder = builder.WithRedirectUri(redirectUrl);
            }
            this.username = username;
            this.password = password;
            publicClientApplication = builder.Build();

            // register tokencache if callback provided
            tokenCacheCallback?.Invoke(publicClientApplication.UserTokenCache);
            authenticationType = ClientContextType.AzureADCredentials;
        }



        /// <summary>
        /// Creates a new instance of the Authentication Manager to acquire access tokens and client contexts using the Azure AD Interactive flow.
        /// </summary>
        /// <param name="clientId">The client id of the Azure AD application to use for authentication</param>
        /// <param name="openBrowserCallback">This callback will be called providing the URL and port to open during the authentication flow</param>
        /// <param name="tenantId">Optional tenant id or tenant url</param>
        /// <param name="successMessageHtml">Allows you to override the success message. Notice that a success header message will be added.</param>
        /// <param name="failureMessageHtml">llows you to override the failure message. Notice that a failed header message will be added and the error message will be appended.</param>
        /// <param name="azureEnvironment">The azure environment to use. Defaults to AzureEnvironment.Production</param>
        /// <param name="tokenCacheCallback">If present, after setting up the base flow for authentication this callback will be called to register a custom tokencache. See https://aka.ms/msal-net-token-cache-serialization.</param>
        public AuthenticationManager(string clientId, Action<string, int> openBrowserCallback, string tenantId = null, string successMessageHtml = null, string failureMessageHtml = null, AzureEnvironment azureEnvironment = AzureEnvironment.Production, Action<ITokenCache> tokenCacheCallback = null) : this(clientId, Utilities.OAuth.DefaultBrowserUi.FindFreeLocalhostRedirectUri(), tenantId, azureEnvironment, tokenCacheCallback, new Utilities.OAuth.DefaultBrowserUi(openBrowserCallback, successMessageHtml, failureMessageHtml))
        {
        }

        /// <summary>
        /// Creates a new instance of the Authentication Manager to acquire access tokens and client contexts using the Azure AD Interactive flow.
        /// </summary>
        /// <param name="clientId">The client id of the Azure AD application to use for authentication</param>
        /// <param name="redirectUrl">Optional redirect URL to use for authentication as set up in the Azure AD Application</param>
        /// <param name="tenantId">Optional tenant id or tenant url</param>
        /// <param name="azureEnvironment">The azure environment to use. Defaults to AzureEnvironment.Production</param>
        /// <param name="tokenCacheCallback">If present, after setting up the base flow for authentication this callback will be called register a custom tokencache. See https://aka.ms/msal-net-token-cache-serialization.</param>
        /// <param name="customWebUi">Optional ICustomWebUi object to fully customize the feedback behavior</param>
        public AuthenticationManager(string clientId, string redirectUrl = null, string tenantId = null, AzureEnvironment azureEnvironment = AzureEnvironment.Production, Action<ITokenCache> tokenCacheCallback = null, ICustomWebUi customWebUi = null) : this()
        {
            this.azureEnvironment = azureEnvironment;

            var builder = PublicClientApplicationBuilder.Create(clientId).WithHttpClientFactory(HttpClientFactory);

            builder = GetBuilderWithAuthority(builder, azureEnvironment);

            if (!string.IsNullOrEmpty(redirectUrl))
            {
                builder = builder.WithRedirectUri(redirectUrl);
            }
            if (!string.IsNullOrEmpty(tenantId))
            {
                builder = builder.WithTenantId(tenantId);
            }
            publicClientApplication = builder.Build();

            this.customWebUi = customWebUi;

            // register tokencache if callback provided
            tokenCacheCallback?.Invoke(publicClientApplication.UserTokenCache);

            authenticationType = ClientContextType.AzureADInteractive;
        }

        /// <summary>
        /// Creates a new instance of the Authentication Manager to acquire authenticated ClientContexts through device code authentication
        /// </summary>
        /// <param name="clientId">The client id of the Azure AD application to use for authentication</param>
        /// <param name="deviceCodeCallback">The callback that will be called with device code information.</param>
        /// <param name="azureEnvironment">The azure environment to use. Defaults to AzureEnvironment.Production</param>
        /// <param name="tokenCacheCallback">If present, after setting up the base flow for authentication this callback will be called register a custom tokencache. See https://aka.ms/msal-net-token-cache-serialization.</param>
        public AuthenticationManager(string clientId, Func<DeviceCodeResult, Task> deviceCodeCallback, AzureEnvironment azureEnvironment = AzureEnvironment.Production, Action<ITokenCache> tokenCacheCallback = null) :
            this(clientId, null, deviceCodeCallback, azureEnvironment, tokenCacheCallback)
        {
        }

        /// <summary>
        /// Creates a new instance of the Authentication Manager to acquire authenticated ClientContexts through device code authentication
        /// </summary>
        /// <param name="clientId">The client id of the Azure AD application to use for authentication</param>
        /// <param name="tenantId">Optional tenant id or tenant url</param>
        /// <param name="deviceCodeCallback">The callback that will be called with device code information.</param>
        /// <param name="azureEnvironment">The azure environment to use. Defaults to AzureEnvironment.Production</param>
        /// <param name="tokenCacheCallback">If present, after setting up the base flow for authentication this callback will be called register a custom tokencache. See https://aka.ms/msal-net-token-cache-serialization.</param>
        public AuthenticationManager(string clientId, string tenantId, Func<DeviceCodeResult, Task> deviceCodeCallback, AzureEnvironment azureEnvironment = AzureEnvironment.Production, Action<ITokenCache> tokenCacheCallback = null) : this()
        {
            this.azureEnvironment = azureEnvironment;
            var azureADEndPoint = GetAzureADLoginEndPoint(azureEnvironment);
            this.deviceCodeCallback = deviceCodeCallback;

            var builder = PublicClientApplicationBuilder.Create(clientId);

            if (!string.IsNullOrEmpty(tenantId))
            {
                builder = builder.WithAuthority($"{azureADEndPoint}/{tenantId}/");
            }
            else
            {
                builder = builder.WithAuthority($"{azureADEndPoint}/organizations/");
            }

            builder = builder.WithHttpClientFactory(HttpClientFactory);

            publicClientApplication = builder.Build();

            // register tokencache if callback provided
            tokenCacheCallback?.Invoke(publicClientApplication.UserTokenCache);

            authenticationType = ClientContextType.DeviceLogin;
        }

        /// <summary>
        /// Creates a new instance of the Authentication Manager to acquire authenticated ClientContexts.
        /// </summary>
        /// <param name="clientId">The client id of the Azure AD application to use for authentication</param>
        /// <param name="certificate">A valid certificate</param>
        /// <param name="tenantId">Tenant id or tenant url</param>
        /// <param name="redirectUrl">Optional redirect URL to use for authentication as set up in the Azure AD Application</param>
        /// <param name="azureEnvironment">The azure environment to use. Defaults to AzureEnvironment.Production</param>
        /// <param name="tokenCacheCallback">If present, after setting up the base flow for authentication this callback will be called register a custom tokencache. See https://aka.ms/msal-net-token-cache-serialization.</param>
        public AuthenticationManager(string clientId, X509Certificate2 certificate, string tenantId, string redirectUrl = null, AzureEnvironment azureEnvironment = AzureEnvironment.Production, Action<ITokenCache> tokenCacheCallback = null) : this()
        {
            this.azureEnvironment = azureEnvironment;
            var azureADEndPoint = GetAzureADLoginEndPoint(azureEnvironment);
            ConfidentialClientApplicationBuilder builder;
            if (azureEnvironment != AzureEnvironment.Production)
            {
                builder = ConfidentialClientApplicationBuilder.Create(clientId).WithCertificate(certificate).WithTenantId(tenantId).WithAuthority(azureADEndPoint, tenantId, true).WithHttpClientFactory(HttpClientFactory);
            }
            else
            {
                builder = ConfidentialClientApplicationBuilder.Create(clientId).WithCertificate(certificate).WithTenantId(tenantId).WithHttpClientFactory(HttpClientFactory);
            }

            if (!string.IsNullOrEmpty(redirectUrl))
            {
                builder = builder.WithRedirectUri(redirectUrl);
            }

            confidentialClientApplication = builder.Build();

            // register tokencache if callback provided
            tokenCacheCallback?.Invoke(confidentialClientApplication.UserTokenCache);

            authenticationType = ClientContextType.AzureADCertificate;
        }

        /// <summary>
        /// Creates a new instance of the Authentication Manager to acquire authenticated ClientContexts.
        /// </summary>
        /// <param name="clientId">The client id of the Azure AD application to use for authentication</param>
        /// <param name="certificatePath">A valid path to a certificate file</param>
        /// <param name="certificatePassword">The password for the certificate</param>
        /// <param name="tenantId">The tenant id (guid) or name (e.g. contoso.onmicrosoft.com) </param>
        /// <param name="redirectUrl">Optional redirect URL to use for authentication as set up in the Azure AD Application</param>
        /// <param name="azureEnvironment">The azure environment to use. Defaults to AzureEnvironment.Production</param>
        /// <param name="tokenCacheCallback">If present, after setting up the base flow for authentication this callback will be called register a custom tokencache. See https://aka.ms/msal-net-token-cache-serialization.</param>
        public AuthenticationManager(string clientId, string certificatePath, string certificatePassword, string tenantId, string redirectUrl = null, AzureEnvironment azureEnvironment = AzureEnvironment.Production, Action<ITokenCache> tokenCacheCallback = null) : this()
        {
            this.azureEnvironment = azureEnvironment;
            var azureADEndPoint = GetAzureADLoginEndPoint(azureEnvironment);

            if (System.IO.File.Exists(certificatePath))
            {
                ConfidentialClientApplicationBuilder builder = null;

                using (var certfile = System.IO.File.OpenRead(certificatePath))
                {
                    var certificateBytes = new byte[certfile.Length];
                    certfile.Read(certificateBytes, 0, (int)certfile.Length);
                    // Don't dispose the cert as that will lead to "m_safeCertContext is an invalid handle" errors when the confidential client actually uses the cert
#pragma warning disable CA2000 // Dispose objects before losing scope
                    var certificate = new X509Certificate2(certificateBytes,
                                                           certificatePassword,
                                                           X509KeyStorageFlags.Exportable |
                                                           X509KeyStorageFlags.MachineKeySet |
                                                           X509KeyStorageFlags.PersistKeySet);
#pragma warning restore CA2000 // Dispose objects before losing scope

                    builder = ConfidentialClientApplicationBuilder.Create(clientId).WithCertificate(certificate).WithTenantId(tenantId).WithHttpClientFactory(HttpClientFactory);
                }

                if (azureEnvironment != AzureEnvironment.Production)
                {
                    builder.WithAuthority(azureADEndPoint, tenantId, true);
                }

                if (!string.IsNullOrEmpty(redirectUrl))
                {
                    builder = builder.WithRedirectUri(redirectUrl);
                }

                confidentialClientApplication = builder.Build();

                // register tokencache if callback provided. ApptokenCache as AcquireTokenForClient is beind called to acquire tokens.
                tokenCacheCallback?.Invoke(confidentialClientApplication.AppTokenCache);

                authenticationType = ClientContextType.AzureADCertificate;

            }
            else
            {
                throw new Exception("Certificate path not found");
            }
        }

        /// <summary>
        /// Creates a new instance of the Authentication Manager to acquire authenticated ClientContexts.
        /// </summary>
        /// <param name="clientId">The client id of the Azure AD application to use for authentication</param>
        /// <param name="storeName">The name of the certificate store to find the certificate in.</param>
        /// <param name="storeLocation">The location of the certificate store to find the certificate in.</param>
        /// <param name="thumbPrint">The thumbprint of the certificate to use.</param>
        /// <param name="tenantId">The tenant id (guid) or name (e.g. contoso.onmicrosoft.com) </param>
        /// <param name="redirectUrl">Optional redirect URL to use for authentication as set up in the Azure AD Application</param>
        /// <param name="azureEnvironment">The azure environment to use. Defaults to AzureEnvironment.Production</param>
        /// <param name="tokenCacheCallback">If present, after setting up the base flow for authentication this callback will be called register a custom tokencache. See https://aka.ms/msal-net-token-cache-serialization.</param>
        public AuthenticationManager(string clientId, StoreName storeName, StoreLocation storeLocation, string thumbPrint, string tenantId, string redirectUrl = null, AzureEnvironment azureEnvironment = AzureEnvironment.Production, Action<ITokenCache> tokenCacheCallback = null) : this()
        {
            this.azureEnvironment = azureEnvironment;
            var azureADEndPoint = GetAzureADLoginEndPoint(azureEnvironment);

            var certificate = Utilities.X509CertificateUtility.LoadCertificate(storeName, storeLocation, thumbPrint);

            var builder = ConfidentialClientApplicationBuilder.Create(clientId).WithCertificate(certificate).WithTenantId(tenantId).WithHttpClientFactory(HttpClientFactory);

            builder = GetBuilderWithAuthority(builder, azureEnvironment, tenantId);

            if (!string.IsNullOrEmpty(redirectUrl))
            {
                builder = builder.WithRedirectUri(redirectUrl);
            }
            if (!string.IsNullOrEmpty(tenantId))
            {
                builder = builder.WithTenantId(tenantId);
            }

            confidentialClientApplication = builder.Build();

            // register tokencache if callback provided. ApptokenCache as AcquireTokenForClient is beind called to acquire tokens.
            tokenCacheCallback?.Invoke(confidentialClientApplication.AppTokenCache);

            authenticationType = ClientContextType.AzureADCertificate;
        }

        /// <summary>
        /// Creates a new instance of the Authentication Manager to acquire authenticated ClientContext.
        /// </summary>
        /// <param name="clientId">The client id of the Azure AD application to use for authentication.</param>
        /// <param name="clientSecret">The client secret of the Azure AD application to use for authentication.</param>
        /// <param name="tenantId">Optional tenant id or tenant url</param>
        /// <param name="azureEnvironment">The azure environment to use. Defaults to AzureEnvironment.Production</param>
        /// <param name="userAssertion">The user assertion (token) of the user on whose behalf to acquire the context</param>
        /// <param name="tokenCacheCallback"></param>
        public AuthenticationManager(string clientId, string clientSecret, UserAssertion userAssertion, string tenantId = null, AzureEnvironment azureEnvironment = AzureEnvironment.Production, Action<ITokenCache> tokenCacheCallback = null) : this()
        {
            this.azureEnvironment = azureEnvironment;
            var azureADEndPoint = GetAzureADLoginEndPoint(azureEnvironment);

            ConfidentialClientApplicationBuilder builder;
            if (azureEnvironment != AzureEnvironment.Production)
            {
                if (tenantId == null)
                {
                    throw new ArgumentException("tenantId is required", nameof(tenantId));
                }
                builder = ConfidentialClientApplicationBuilder.Create(clientId).WithClientSecret(clientSecret).WithAuthority(azureADEndPoint, tenantId, true).WithHttpClientFactory(HttpClientFactory);
            }
            else
            {
                builder = ConfidentialClientApplicationBuilder.Create(clientId).WithClientSecret(clientSecret).WithAuthority($"{azureADEndPoint}/organizations/").WithHttpClientFactory(HttpClientFactory);
                if (!string.IsNullOrEmpty(tenantId))
                {
                    builder = builder.WithTenantId(tenantId);
                }
            }
            this.assertion = userAssertion;
            confidentialClientApplication = builder.Build();

            // register tokencache if callback provided
            tokenCacheCallback?.Invoke(confidentialClientApplication.UserTokenCache);
            authenticationType = ClientContextType.AzureOnBehalfOf;
        }

        /// <summary>
        /// Creates an AuthenticationManager for the given PnP Core SDK <see cref="IAuthenticationProvider"/>.
        /// </summary>
        /// <param name="authenticationProvider">PnP Core SDK <see cref="IAuthenticationProvider"/></param>
        public AuthenticationManager(IAuthenticationProvider authenticationProvider)
        {
            this.authenticationProvider = authenticationProvider;
            this.pnpContext = null;
            authenticationType = ClientContextType.PnPCoreSdk;
        }

        /// <summary>
        /// Creates an AuthenticationManager for the given PnP Core SDK
        /// </summary>
        /// <param name="pnPContext">PnP Core SDK<see cref="PnPContext"/></param>
        public AuthenticationManager(PnPContext pnPContext)
        {
            this.authenticationProvider = pnPContext.AuthenticationProvider;
            this.pnpContext = pnPContext;
            authenticationType = ClientContextType.PnPCoreSdk;
            ConfigureAuthenticationManagerEnvironmentSettings(pnPContext);
        }

        private void ConfigureAuthenticationManagerEnvironmentSettings(PnPContext pnPContext)
        {
            if (pnPContext.Environment == Microsoft365Environment.Custom)
            {
                this.azureEnvironment = AzureEnvironment.Custom;
                this.microsoftGraphEndPoint = pnPContext.MicrosoftGraphAuthority;
                this.azureADLoginEndPoint = $"https://{pnPContext.AzureADLoginAuthority}";
            }
            else
            {
                this.azureEnvironment = pnPContext.Environment switch
                {
                    Microsoft365Environment.Production => AzureEnvironment.Production,
                    Microsoft365Environment.Germany => AzureEnvironment.Germany,
                    Microsoft365Environment.China => AzureEnvironment.China,
                    Microsoft365Environment.USGovernment => AzureEnvironment.USGovernment,
                    Microsoft365Environment.USGovernmentHigh => AzureEnvironment.USGovernmentHigh,
                    Microsoft365Environment.USGovernmentDoD => AzureEnvironment.USGovernmentDoD,
                    Microsoft365Environment.PreProduction => AzureEnvironment.PPE,
                    _ => AzureEnvironment.Production
                };
            }
        }
        #endregion

        #region Access Token Acquisition
        /// <summary>
        /// Returns an access token for a given site.
        /// </summary>
        /// <param name="siteUrl"></param>
        /// <param name="prompt">The prompt style to use. Notice that this only works with the Interactive Login flow, for all other flows this parameter is ignored.</param>
        /// <returns></returns>
        public string GetAccessToken(string siteUrl, Prompt prompt = default)
        {
            return GetAccessTokenAsync(siteUrl, CancellationToken.None, prompt).GetAwaiter().GetResult();
        }

        /// <summary>
        /// Returns an access token for a given site.
        /// </summary>
        /// <param name="siteUrl"></param>
        /// <param name="prompt">The prompt style to use. Notice that this only works with the Interactive Login flow, for all other flows this parameter is ignored.</param>
        /// <returns></returns>
        public async Task<string> GetAccessTokenAsync(string siteUrl, Prompt prompt = default)
        {
            return await GetAccessTokenAsync(siteUrl, CancellationToken.None, prompt).ConfigureAwait(false);
        }

        /// <summary>
        /// Returns an access token for a given site.
        /// </summary>
        /// <param name="siteUrl"></param>
        /// <param name="cancellationToken">Optional cancellation token to cancel the request</param>
        /// <param name="prompt">The prompt style to use. Notice that this only works with the Interactive Login flow, for all other flows this parameter is ignored.</param>
        /// <returns></returns>
        public string GetAccessToken(string siteUrl, CancellationToken cancellationToken, Prompt prompt = default)
        {
            var uri = new Uri(siteUrl);

            var scopes = new[] { $"{uri.Scheme}://{uri.Authority}/.default" };

            return GetAccessTokenAsync(scopes, cancellationToken, prompt, uri).GetAwaiter().GetResult();
        }

        /// <summary>
        /// Returns an access token for a given site.
        /// </summary>
        /// <param name="siteUrl"></param>
        /// <param name="cancellationToken">Optional cancellation token to cancel the request</param>
        /// <param name="prompt">The prompt style to use. Notice that this only works with the Interactive Login flow, for all other flows this parameter is ignored.</param>
        /// <returns></returns>
        public async Task<string> GetAccessTokenAsync(string siteUrl, CancellationToken cancellationToken, Prompt prompt = default)
        {
            var uri = new Uri(siteUrl);

            var scopes = new[] { $"{uri.Scheme}://{uri.Authority}/.default" };

            return await GetAccessTokenAsync(scopes, cancellationToken, prompt, uri).ConfigureAwait(false);
        }

        /// <summary>
        /// Returns an access token for the given scopes.
        /// </summary>
        /// <param name="scopes">The scopes to retrieve the access token for</param>
        /// <param name="prompt">The prompt style to use. Notice that this only works with the Interactive Login flow, for all other flows this parameter is ignored.</param>
        /// <returns></returns>
        public async Task<string> GetAccessTokenAsync(string[] scopes, Prompt prompt = default)
        {
            return await GetAccessTokenAsync(scopes, CancellationToken.None, prompt).ConfigureAwait(false);
        }


        /// <summary>
        /// Returns an access token for the given scopes.
        /// </summary>
        /// <param name="scopes">The scopes to retrieve the access token for</param>
        /// <param name="cancellationToken">Optional cancellation token to cancel the request</param>
        /// <param name="prompt">The prompt style to use. Notice that this only works with the Interactive Login flow, for all other flows this parameter is ignored.</param>
        /// <param name="uri">for ClientContextType.PnPCoreSdk case as by interface definition needed for GetAccessTokenAsync</param>
        /// <returns></returns>
        public async Task<string> GetAccessTokenAsync(string[] scopes, CancellationToken cancellationToken, Prompt prompt = default, Uri uri = null)
        {
            AuthenticationResult authResult = null;

            switch (authenticationType)
            {
                case ClientContextType.AzureADCredentials:
                    {
                        var accounts = await publicClientApplication.GetAccountsAsync().ConfigureAwait(false);
                        try
                        {
                            authResult = await publicClientApplication.AcquireTokenSilent(scopes, accounts.First()).ExecuteAsync(cancellationToken).ConfigureAwait(false);
                        }
                        catch
                        {
#pragma warning disable CS0618 // Type or member is obsolete
                            authResult = await publicClientApplication.AcquireTokenByUsernamePassword(scopes, username, password).ExecuteAsync(cancellationToken).ConfigureAwait(false);
#pragma warning restore CS0618 // Type or member is obsolete
                        }
                        break;
                    }
                case ClientContextType.AzureADInteractive:
                    {
                        var accounts = await publicClientApplication.GetAccountsAsync().ConfigureAwait(false);

                        try
                        {
                            authResult = await publicClientApplication.AcquireTokenSilent(scopes, accounts.First()).ExecuteAsync(cancellationToken).ConfigureAwait(false);
                        }
                        catch
                        {
                            var builder = publicClientApplication.AcquireTokenInteractive(scopes);
                            if (RuntimeInformation.IsOSPlatform(OSPlatform.OSX))
                            {
                                var options = new SystemWebViewOptions()
                                {
                                    HtmlMessageError = "<p> An error occurred: {0}. Details {1}</p>",
                                    HtmlMessageSuccess = "<p>Succesfully acquired token. You may close this window now.</p>"
                                };
                                builder = builder.WithUseEmbeddedWebView(false);
                                builder = builder.WithSystemWebViewOptions(options);
                            }
                            else
                            {
                                if (customWebUi != null)
                                {
                                    builder = builder.WithCustomWebUi(customWebUi);
                                }
                                if (prompt != default)
                                {
                                    builder.WithPrompt(prompt);
                                }
                            }
                            authResult = await builder.ExecuteAsync(cancellationToken).ConfigureAwait(false);
                        }
                        break;
                    }
                case ClientContextType.AzureADCertificate:
                    {
#pragma warning disable CS0618 // Type or member is obsolete
                        var accounts = await confidentialClientApplication.GetAccountsAsync().ConfigureAwait(false);
#pragma warning restore CS0618 // Type or member is obsolete

                        try
                        {
                            authResult = await confidentialClientApplication.AcquireTokenSilent(scopes, accounts.First()).ExecuteAsync(cancellationToken).ConfigureAwait(false);
                        }
                        catch
                        {
                            var builder = confidentialClientApplication.AcquireTokenForClient(scopes);

                            authResult = await confidentialClientApplication.AcquireTokenForClient(scopes).ExecuteAsync(cancellationToken).ConfigureAwait(false);
                        }
                        break;
                    }
                case ClientContextType.DeviceLogin:
                    {
                        var accounts = await publicClientApplication.GetAccountsAsync().ConfigureAwait(false);
                        try
                        {
                            authResult = await publicClientApplication.AcquireTokenSilent(scopes, accounts.First()).ExecuteAsync(cancellationToken).ConfigureAwait(false);
                        }
                        catch
                        {
                            authResult = await publicClientApplication.AcquireTokenWithDeviceCode(scopes, deviceCodeCallback).ExecuteAsync(cancellationToken).ConfigureAwait(false);
                        }
                        break;
                    }
                case ClientContextType.AzureOnBehalfOf:
                    {
#pragma warning disable CS0618 // Type or member is obsolete
                        var accounts = await confidentialClientApplication.GetAccountsAsync().ConfigureAwait(false);
#pragma warning restore CS0618 // Type or member is obsolete

                        try
                        {
                            authResult = await confidentialClientApplication.AcquireTokenSilent(scopes, accounts.First()).ExecuteAsync(cancellationToken).ConfigureAwait(false);
                        }
                        catch
                        {
                            authResult = await confidentialClientApplication.AcquireTokenOnBehalfOf(scopes, assertion).ExecuteAsync(cancellationToken).ConfigureAwait(false);
                        }
                        break;
                    }
                case ClientContextType.SharePointACSAppOnly:
                    {
                        if (acsTokenGenerator == null)
                        {
                            throw new ArgumentException($"{nameof(GetAccessTokenAsync)}() called without an ACS token generator. Specify in {nameof(AuthenticationManager)} constructor the authentication parameters");
                        }
                        return acsTokenGenerator.GetToken(null);
                    }
                case ClientContextType.AccessToken:
                    {
                        return new NetworkCredential("", accessToken).Password;
                    }
                case ClientContextType.PnPCoreSdk:
                    {
                        return await this.authenticationProvider.GetAccessTokenAsync(uri, scopes).ConfigureAwait(false);
                    }
            }
            if (authResult?.AccessToken != null)
            {
                return authResult.AccessToken;
            }
            return null;
        }
        #endregion

        #region Context Acquisition
        /// <summary>
        /// Returns a CSOM ClientContext which has been set up for Azure AD OAuth authentication
        /// </summary>
        /// <param name="siteUrl"></param>
        /// <returns></returns>
        public ClientContext GetContext(string siteUrl)
        {
            return GetContextAsync(siteUrl, CancellationToken.None).GetAwaiter().GetResult();
        }

        /// <summary>
        /// Returns a CSOM ClientContext which has been set up for Azure AD OAuth authentication
        /// </summary>
        /// <param name="siteUrl"></param>
        /// <param name="cancellationToken">Optional cancellation token to cancel the request</param>
        /// <returns></returns>
        public ClientContext GetContext(string siteUrl, CancellationToken cancellationToken)
        {
            return GetContextAsync(siteUrl, cancellationToken).GetAwaiter().GetResult();
        }

        /// <summary>
        /// Returns a CSOM ClientContext which has been set up for Azure AD OAuth authentication
        /// </summary>
        /// <param name="siteUrl"></param>
        /// <returns></returns>
        public async Task<ClientContext> GetContextAsync(string siteUrl)
        {
            return await GetContextAsync(siteUrl, CancellationToken.None).ConfigureAwait(false);
        }

        /// <summary>
        /// Returns a CSOM ClientContext which has been set up for Azure AD OAuth authentication
        /// </summary>
        /// <param name="siteUrl"></param>
        /// <param name="cancellationToken">Optional cancellation token to cancel the request</param>
        /// <returns></returns>
        public async Task<ClientContext> GetContextAsync(string siteUrl, CancellationToken cancellationToken)
        {
            var uri = new Uri(siteUrl);

            var scopes = new[] { $"{uri.Scheme}://{uri.Authority}/.default" };

            AuthenticationResult authResult;

            switch (authenticationType)
            {
                case ClientContextType.AzureADCredentials:
                    {
                        var accounts = await publicClientApplication.GetAccountsAsync().ConfigureAwait(false);
                        try
                        {
                            authResult = await publicClientApplication.AcquireTokenSilent(scopes, accounts.First()).ExecuteAsync(cancellationToken).ConfigureAwait(false);
                        }
                        catch
                        {
#pragma warning disable CS0618 // Type or member is obsolete
                            authResult = await publicClientApplication.AcquireTokenByUsernamePassword(scopes, username, password).ExecuteAsync(cancellationToken).ConfigureAwait(false);
#pragma warning restore CS0618 // Type or member is obsolete
                        }
                        if (authResult.AccessToken != null)
                        {
                            return BuildClientContext(publicClientApplication, siteUrl, scopes, authenticationType);
                        }
                        break;
                    }
                case ClientContextType.AzureADInteractive:
                    {
                        var accounts = await publicClientApplication.GetAccountsAsync().ConfigureAwait(false);

                        try
                        {
                            authResult = await publicClientApplication.AcquireTokenSilent(scopes, accounts.First()).ExecuteAsync(cancellationToken).ConfigureAwait(false);
                        }
                        catch
                        {
                            var builder = publicClientApplication.AcquireTokenInteractive(scopes);
                            if (RuntimeInformation.IsOSPlatform(OSPlatform.OSX))
                            {
                                var options = new SystemWebViewOptions()
                                {
                                    HtmlMessageError = "<p> An error occurred: {0}. Details {1}</p>",
                                    HtmlMessageSuccess = "<p>Succesfully acquired token. You may close this window now.</p>"
                                };
                                builder = builder.WithUseEmbeddedWebView(false);
                                builder = builder.WithSystemWebViewOptions(options);
                            }
                            else
                            {

                                if (customWebUi != null)
                                {
                                    builder = builder.WithCustomWebUi(customWebUi);
                                }
                            }
                            authResult = await builder.ExecuteAsync(cancellationToken).ConfigureAwait(false);
                        }
                        if (authResult.AccessToken != null)
                        {
                            return BuildClientContext(publicClientApplication, siteUrl, scopes, authenticationType);
                        }
                        break;
                    }
                case ClientContextType.AzureADCertificate:
                    {
#pragma warning disable CS0618 // Type or member is obsolete
                        var accounts = await confidentialClientApplication.GetAccountsAsync().ConfigureAwait(false);
#pragma warning restore CS0618 // Type or member is obsolete

                        try
                        {
                            authResult = await confidentialClientApplication.AcquireTokenSilent(scopes, accounts.First()).ExecuteAsync(cancellationToken).ConfigureAwait(false);
                        }
                        catch
                        {
                            authResult = await confidentialClientApplication.AcquireTokenForClient(scopes).ExecuteAsync(cancellationToken).ConfigureAwait(false);
                        }
                        if (authResult.AccessToken != null)
                        {
                            return BuildClientContext(confidentialClientApplication, siteUrl, scopes, authenticationType);
                        }
                        break;
                    }
                case ClientContextType.AzureOnBehalfOf:
                    {
#pragma warning disable CS0618 // Type or member is obsolete
                        var accounts = await confidentialClientApplication.GetAccountsAsync().ConfigureAwait(false);
#pragma warning restore CS0618 // Type or member is obsolete

                        try
                        {
                            authResult = await confidentialClientApplication.AcquireTokenSilent(scopes, accounts.First()).ExecuteAsync(cancellationToken).ConfigureAwait(false);
                        }
                        catch
                        {
                            authResult = await confidentialClientApplication.AcquireTokenOnBehalfOf(scopes, assertion).ExecuteAsync(cancellationToken).ConfigureAwait(false);
                        }
                        if (authResult.AccessToken != null)
                        {
                            return BuildClientContext(confidentialClientApplication, siteUrl, scopes, authenticationType);
                        }
                        break;
                    }
                case ClientContextType.DeviceLogin:
                    {
                        var accounts = await publicClientApplication.GetAccountsAsync().ConfigureAwait(false);

                        try
                        {
                            authResult = await publicClientApplication.AcquireTokenSilent(scopes, accounts.First()).ExecuteAsync(cancellationToken).ConfigureAwait(false);
                        }
                        catch
                        {
                            authResult = await publicClientApplication.AcquireTokenWithDeviceCode(scopes, deviceCodeCallback).ExecuteAsync(cancellationToken).ConfigureAwait(false);
                        }
                        if (authResult.AccessToken != null)
                        {
                            return BuildClientContext(publicClientApplication, siteUrl, scopes, authenticationType);
                        }
                        break;
                    }

                case ClientContextType.SharePointACSAppOnly:
                    {
                        if (acsTokenGenerator == null)
                        {
                            throw new ArgumentException($"{nameof(GetContextAsync)}() called without an ACS token generator. Use {nameof(GetACSAppOnlyContext)}() or {nameof(GetAccessTokenContext)}() instead or specify in {nameof(AuthenticationManager)} constructor the authentication parameters");
                        }

                        var context = GetAccessTokenContext(siteUrl, (site) =>
                        {
                            return acsTokenGenerator.GetToken(new Uri(site));
                        });

                        ClientContextSettings clientContextSettings = new ClientContextSettings()
                        {
                            Type = ClientContextType.SharePointACSAppOnly,
                            SiteUrl = siteUrl,
                            AuthenticationManager = this,
                            Environment = this.azureEnvironment
                        };
                        context.AddContextSettings(clientContextSettings);

                        return context;

                    }

                case ClientContextType.AccessToken:
                    {
                        var context = GetAccessTokenContext(siteUrl, (site) =>
                        {
                            return EncryptionUtility.ToInsecureString(this.accessToken);
                        });
                        ClientContextSettings clientContextSettings = new ClientContextSettings()
                        {
                            Type = ClientContextType.AccessToken,
                            SiteUrl = siteUrl,
                            AuthenticationManager = this,
                            Environment = this.azureEnvironment
                        };
                        context.AddContextSettings(clientContextSettings);

                        return context;
                    }
                case ClientContextType.PnPCoreSdk:
                    {
                        if (authenticationProvider == null)
                        {
                            throw new ArgumentException($"{nameof(GetContextAsync)}() called without an IAuthenticationProvider.");
                        }

                        var context = GetAccessTokenContext(siteUrl, (site) =>
                        {
                            return authenticationProvider.GetAccessTokenAsync(new Uri(site)).GetAwaiter().GetResult();
                        });

                        ClientContextSettings clientContextSettings = new ClientContextSettings()
                        {
                            Type = ClientContextType.PnPCoreSdk,
                            SiteUrl = siteUrl,
                            AuthenticationManager = this,
                            Environment = this.azureEnvironment
                        };
                        context.AddContextSettings(clientContextSettings);

                        return context;
                    }
            }
            return null;
        }

        /// <summary>
        /// Return same IAuthenticationProvider then the AuthenticationManager was initialized with
        /// </summary>
        internal IAuthenticationProvider PnPCoreAuthenticationProvider
        {
            get
            {
                if (authenticationType == ClientContextType.PnPCoreSdk && authenticationProvider != null)
                {
                    return authenticationProvider;
                }
                else
                {
                    return null;
                }
            }
        }

        /// <summary>
        /// Return same PnPContext then the AuthenticationManager was initialized with
        /// </summary>
        internal PnPContext PnPCoreContext
        {
            get
            {
                if (authenticationType == ClientContextType.PnPCoreSdk && pnpContext != null)
                {
                    return pnpContext;
                }
                else
                {
                    return null;
                }
            }
        }
        #endregion

        #region Internals
        private ClientContext BuildClientContext(IClientApplicationBase application, string siteUrl, string[] scopes, ClientContextType contextType)
        {
            var clientContext = new ClientContext(siteUrl)
            {
                DisableReturnValueCache = true
            };

            clientContext.ExecutingWebRequest += (sender, args) =>
            {
                AuthenticationResult ar = null;

                var accounts = application.GetAccountsAsync().GetAwaiter().GetResult();
                if (accounts.Any())
                {
                    ar = application.AcquireTokenSilent(scopes, accounts.First()).ExecuteAsync().GetAwaiter().GetResult();
                }
                else
                {
                    switch (contextType)
                    {
                        case ClientContextType.AzureADCertificate:
                            {
                                ar = ((IConfidentialClientApplication)application).AcquireTokenForClient(scopes).ExecuteAsync().GetAwaiter().GetResult();
                                break;
                            }
                        case ClientContextType.AzureADCredentials:
                            {
#pragma warning disable CS0618 // Type or member is obsolete
                                ar = ((IPublicClientApplication)application).AcquireTokenByUsernamePassword(scopes, username, password).ExecuteAsync().GetAwaiter().GetResult();
#pragma warning restore CS0618 // Type or member is obsolete
                                break;
                            }
                        case ClientContextType.AzureADInteractive:
                            {
                                var builder = ((IPublicClientApplication)application).AcquireTokenInteractive(scopes);
                                if (customWebUi != null)
                                {
                                    builder = builder.WithCustomWebUi(customWebUi);
                                }
                                ar = builder.ExecuteAsync().GetAwaiter().GetResult();
                                break;
                            }
                        case ClientContextType.AzureOnBehalfOf:
                            {
                                ar = ((IConfidentialClientApplication)application).AcquireTokenOnBehalfOf(scopes, assertion).ExecuteAsync().GetAwaiter().GetResult();
                                break;
                            }
                        case ClientContextType.DeviceLogin:
                            {
                                ar = ((IPublicClientApplication)application).AcquireTokenWithDeviceCode(scopes, deviceCodeCallback).ExecuteAsync().GetAwaiter().GetResult();
                                break;
                            }
                    }
                }
                if (ar != null && ar.AccessToken != null)
                {
                    args.WebRequestExecutor.RequestHeaders["Authorization"] = "Bearer " + ar.AccessToken;
                }
            };

            ClientContextSettings clientContextSettings = new ClientContextSettings()
            {
                Type = contextType,
                SiteUrl = siteUrl,
                AuthenticationManager = this,
                Environment = this.azureEnvironment
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

        public ClientContext GetOnPremisesContext(string siteUrl, string userName, SecureString password)
        {
            ClientContext clientContext = new ClientContext(siteUrl)
            {
                DisableReturnValueCache = true,
                Credentials = new NetworkCredential(userName, password)
            };

            ConfigureOnPremisesContext(siteUrl, clientContext);

            return clientContext;
        }

        public ClientContext GetOnPremisesContext(string siteUrl, ICredentials credentials)
        {
            ClientContext clientContext = new ClientContext(siteUrl)
            {
                DisableReturnValueCache = true,
                Credentials = credentials
            };

            ConfigureOnPremisesContext(siteUrl, clientContext);

            return clientContext;
        }

        public ClientContext GetOnPremisesContext(string siteUrl)
        {
            ClientContext clientContext = new ClientContext(siteUrl)
            {
                DisableReturnValueCache = true,
                Credentials = CredentialCache.DefaultNetworkCredentials
            };

            ConfigureOnPremisesContext(siteUrl, clientContext);

            return clientContext;
        }

        internal void ConfigureOnPremisesContext(string siteUrl, ClientContext clientContext)
        {
            clientContext.ExecutingWebRequest += (sender, webRequestEventArgs) =>
            {
                // CSOM for .NET Standard 2.0 is not sending along credentials for an on-premises request, so ensure 
                // credentials and request digest are in place. This will make CSOM for .NET Standard work for 
                // SharePoint 2013, 2016 and 2019. For SharePoint 2010 this does not work as the generated CSOM request
                // contains references to version 15 while 2010 expects version 14.
                //
                // Note: the "onpremises" part of AuthenticationManager internal by design as it's only intended to be
                //       used by transformation tech that needs to get data from on-premises. PnP Framework, nor PnP 
                //       PowerShell do support SharePoint on-premises.
                webRequestEventArgs.WebRequestExecutor.WebRequest.Credentials = (sender as ClientContext).Credentials;
                // CSOM for .NET Standard does not handle request digest management, a POST to client.svc requires a digest, so ensuring that
                webRequestEventArgs.WebRequestExecutor.WebRequest.Headers.Add("X-RequestDigest", (sender as ClientContext).GetOnPremisesRequestDigestAsync().GetAwaiter().GetResult());
                // Add Request Header to force Windows Authentication which avoids an issue if multiple authentication providers are enabled on a webapplication
                webRequestEventArgs.WebRequestExecutor.RequestHeaders["X-FORMS_BASED_AUTH_ACCEPTED"] = "f";
            };

            ClientContextSettings clientContextSettings = new ClientContextSettings()
            {
                Type = ClientContextType.OnPremises,
                SiteUrl = siteUrl,
                AuthenticationManager = this
            };

            clientContext.AddContextSettings(clientContextSettings);
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
            return this.GetACSAppOnlyContext(siteUrl, appId, appSecret, AzureEnvironment.Production);
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
            return this.GetACSAppOnlyContext(siteUrl, null, appId, appSecret, GetACSEndPoint(environment), GetACSEndPointPrefix(environment));
        }

        /// <summary>
        /// Returns an app only ClientContext object
        /// </summary>
        /// <param name="siteUrl">Site for which the ClientContext object will be instantiated</param>
        /// <param name="realm">Realm of the environment (tenant) that requests the ClientContext object, may be null</param>
        /// <param name="appId">Application ID which is requesting the ClientContext object</param>
        /// <param name="appSecret">Application secret of the Application which is requesting the ClientContext object</param>
        /// <param name="acsHostUrl">Azure ACS host, defaults to accesscontrol.windows.net but internal pre-production environments use other hosts</param>
        /// <param name="globalEndPointPrefix">Azure ACS endpoint prefix, defaults to accounts but internal pre-production environments use other prefixes</param>
        /// <returns>ClientContext to be used by CSOM code</returns>
        public ClientContext GetACSAppOnlyContext(string siteUrl, string realm, string appId, string appSecret, string acsHostUrl = "accesscontrol.windows.net", string globalEndPointPrefix = "accounts")
        {
            var acsTokenProvider = ACSTokenGenerator.GetACSAuthenticationProvider(new Uri(siteUrl), realm, appId, appSecret, acsHostUrl, globalEndPointPrefix);
            var am = new AuthenticationManager(acsTokenProvider);
            ClientContext clientContext = am.GetContext(siteUrl);

            ClientContextSettings clientContextSettings = new ClientContextSettings()
            {
                Type = ClientContextType.SharePointACSAppOnly,
                SiteUrl = siteUrl,
                AuthenticationManager = am,
                Realm = realm,
                ClientId = appId,
                ClientSecret = appSecret,
                AcsHostUrl = acsHostUrl,
                GlobalEndPointPrefix = globalEndPointPrefix,
                Environment = this.azureEnvironment
            };

            clientContext.AddContextSettings(clientContextSettings);

            return clientContext;
        }

        /// <summary>
        /// Gets the Azure ASC login end point for the given environment
        /// </summary>
        /// <param name="environment">Environment to get the login information for</param>
        /// <returns>Azure ASC login endpoint</returns>
        public static string GetACSEndPoint(AzureEnvironment environment)
        {
            return (environment) switch
            {
                AzureEnvironment.Production => "accesscontrol.windows.net",
                AzureEnvironment.Germany => "microsoftonline.de",
                AzureEnvironment.China => "accesscontrol.chinacloudapi.cn",
                AzureEnvironment.USGovernment => "microsoftonline.us",
                AzureEnvironment.PPE => "windows-ppe.net",
                _ => "accesscontrol.windows.net"
            };
        }

        /// <summary>
        /// Gets the Azure ACS login end point prefix for the given environment
        /// </summary>
        /// <param name="environment">Environment to get the login information for</param>
        /// <returns>Azure ACS login endpoint prefix</returns>
        public static string GetACSEndPointPrefix(AzureEnvironment environment)
        {
            return (environment) switch
            {
                AzureEnvironment.Production => "accounts",
                AzureEnvironment.Germany => "login",
                AzureEnvironment.China => "accounts",
                AzureEnvironment.USGovernment => "login",
                AzureEnvironment.USGovernmentHigh => "login",
                AzureEnvironment.USGovernmentDoD => "login",
                AzureEnvironment.PPE => "login",
                _ => "accounts"
            };
        }


        /// <summary>
        /// Returns a SharePoint ClientContext using a custom access token function. The function will be called with the Resource Uri and expected to return an access token as a string.
        /// </summary>
        /// <param name="siteUrl">Site for which the ClientContext object will be instantiated</param>
        /// <param name="accessTokenGetter">The AccessToken getter method to use</param>
        /// <returns>Client context object</returns>
        public ClientContext GetAccessTokenContext(string siteUrl, Func<string, string> accessTokenGetter)
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
        /// Returns a SharePoint ClientContext using custom provided access token.
        /// </summary>
        /// <param name="siteUrl">Site for which the ClientContext object will be instantiated</param>
        /// <param name="accessToken">An explicit value for the AccessToken</param>
        /// <returns>Client context object</returns>
        public ClientContext GetAccessTokenContext(string siteUrl, string accessToken)
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
        /// Gets the Azure AD login end point for the given environment
        /// </summary>
        /// <param name="environment">Environment to get the login information for</param>
        /// <returns>Azure AD login endpoint</returns>
        public string GetAzureADLoginEndPoint(AzureEnvironment environment)
        {
            if (environment == AzureEnvironment.Custom)
            {
                return GetAzureAdLoginEndPointForCustomAzureEnvironmentConfiguration();
            }
            else
            {
                return GetAzureADLoginEndPointStatic(environment);
            }
        }

        public static string GetAzureADLoginEndPointStatic(AzureEnvironment environment)
        {
            return (environment) switch
            {
                AzureEnvironment.Production => "https://login.microsoftonline.com",
                AzureEnvironment.Germany => "https://login.microsoftonline.de",
                AzureEnvironment.China => "https://login.chinacloudapi.cn",
                AzureEnvironment.USGovernment => "https://login.microsoftonline.us",
                AzureEnvironment.USGovernmentHigh => "https://login.microsoftonline.us",
                AzureEnvironment.USGovernmentDoD => "https://login.microsoftonline.us",
                AzureEnvironment.PPE => "https://login.windows-ppe.net",
                _ => "https://login.microsoftonline.com"
            };
        }

        /// <summary>
        /// Returns the Graph End Point url without protocol based upon the Azure Environment selected during creation of the Authentication Manager
        /// </summary>
        /// <returns></returns>
        public string GetGraphEndPoint()
        {
            if (this.azureEnvironment == AzureEnvironment.Custom)
            {
                return GetGraphEndPointForCustomAzureEnvironmentConfiguration();
            }
            else
            {
                return GetGraphEndPoint(this.azureEnvironment);
            }
        }

        /// <summary>
        /// Returns the Graph End Point url without protocol based upon the provided Azure Environment
        /// </summary>
        /// <returns></returns>
        public static string GetGraphEndPoint(AzureEnvironment environment)
        {
            switch (environment)
            {
                case AzureEnvironment.Production:
                case AzureEnvironment.USGovernment:
                    {
                        return "graph.microsoft.com";
                    }
                case AzureEnvironment.Germany:
                    {
                        return "graph.microsoft.de";
                    }
                case AzureEnvironment.China:
                    {
                        return "microsoftgraph.chinacloudapi.cn";
                    }
                case AzureEnvironment.USGovernmentHigh:
                    {
                        return "graph.microsoft.us";
                    }
                case AzureEnvironment.USGovernmentDoD:
                    {
                        return "dod-graph.microsoft.us";
                    }
                default:
                    {
                        return "graph.microsoft.com";
                    }
            }
        }

        /// <summary>
        /// Gets the URI to use for making Graph calls based upon the environment
        /// </summary>
        /// <param name="environment">Environment to get the Graph URI for</param>
        /// <returns>Graph URI for given environment</returns>
        public static Uri GetGraphBaseEndPoint(AzureEnvironment environment)
        {
            return new Uri($"https://{GetGraphEndPoint(environment)}");
        }

        /// <summary>
        /// Returns a domain suffix (com, us, de, cn) for an Azure Environment
        /// </summary>
        /// <param name="environment"></param>
        /// <returns></returns>
        public static string GetSharePointDomainSuffix(AzureEnvironment environment)
        {
            return (environment) switch
            {
                AzureEnvironment.Production => "com",
                AzureEnvironment.USGovernment => "us",
                AzureEnvironment.USGovernmentHigh => "us",
                AzureEnvironment.USGovernmentDoD => "us",
                AzureEnvironment.Germany => "de",
                AzureEnvironment.China => "cn",
                _ => "com"
            };
        }

        /// <summary>
        /// Returns the equivalent SharePoint Admin url for the passed in SharePoint url
        /// </summary>
        /// <param name="url">Any SharePoint url for the tenant you need to SharePoint Admin Center URL for</param>
        /// <returns>SharePoint Admin Center URL</returns>
        public static string GetTenantAdministrationUrl(string url)
        {
            var uri = new Uri(url);
            var uriParts = uri.Host.Split('.');

            if (uriParts[0].EndsWith("-admin"))
            {
                // The url was already an admin site url 
                return $"https://{uriParts[0]}.{string.Join(".", uriParts.Skip(1))}";
            }

            if (!uriParts[0].EndsWith("-admin"))
            {
                return $"https://{uriParts[0]}-admin.{string.Join(".", uriParts.Skip(1))}";
            }
            return null;
        }

        /// <summary>
        /// Returns the equivalent SharePoint Admin url for the passed in SharePoint url
        /// </summary>
        /// <param name="url">Any SharePoint url for the tenant you need to SharePoint Admin Center URL for</param>
        /// <returns>SharePoint Admin Center URL</returns>
        public static Uri GetTenantAdministrationUri(string url)
        {
            string adminUrl = GetTenantAdministrationUrl(url);
            if (adminUrl != null)
            {
                return new Uri(adminUrl);
            }

            return null;
        }

        /// <summary>
        /// Is the provided URL an SharePoint Admin center URL
        /// </summary>
        /// <param name="url">SharePoint URL to check</param>
        /// <returns>True if Admin Center URL, false otherwise</returns>
        public static bool IsTenantAdministrationUrl(string url)
        {
            return url.ToLowerInvariant().Contains("-admin.sharepoint");
        }

        /// <summary>
        /// Is the provided URL an SharePoint Admin center URL
        /// </summary>
        /// <param name="uri">SharePoint URL to check</param>
        /// <returns>True if Admin Center URL, false otherwise</returns>
        public static bool IsTenantAdministrationUri(Uri uri)
        {
            return IsTenantAdministrationUrl(uri.ToString());
        }

        public string GetGraphEndPointForCustomAzureEnvironmentConfiguration()
        {
            if (string.IsNullOrEmpty(microsoftGraphEndPoint))
            {
                microsoftGraphEndPoint = LoadConfiguration("MicrosoftGraphEndPoint");
            }

            if (string.IsNullOrEmpty(microsoftGraphEndPoint))
            {
                return "graph.microsoft.com";
            }
            else
            {
                return microsoftGraphEndPoint;
            }
        }

        public string GetAzureAdLoginEndPointForCustomAzureEnvironmentConfiguration()
        {
            if (string.IsNullOrEmpty(azureADLoginEndPoint))
            {
                azureADLoginEndPoint = LoadConfiguration("AzureADLoginEndPoint");
            }

            if (string.IsNullOrEmpty(azureADLoginEndPoint))
            {
                return "https://login.microsoftonline.com";
            }
            else
            {
                return azureADLoginEndPoint;
            }
        }

        public void SetEndPointsForCustomAzureEnvironmentConfiguration(string microsoftGraphEndPoint, string azureADLoginEndPoint)
        {
            this.microsoftGraphEndPoint = microsoftGraphEndPoint;
            this.azureADLoginEndPoint = azureADLoginEndPoint;
        }

        private static string LoadConfiguration(string appSetting)
        {
            string loadedAppSetting = null;
            try
            {
                loadedAppSetting = ConfigurationManager.AppSettings[appSetting];
            }
            catch // throws exception if being called from a .NET Standard 2.0 application
            {

            }

            if (string.IsNullOrWhiteSpace(loadedAppSetting))
            {
                loadedAppSetting = Environment.GetEnvironmentVariable(appSetting, EnvironmentVariableTarget.Process);
            }

            return loadedAppSetting;
        }

        /// <summary>
        /// Clears the internal in-memory token cache used by MSAL
        /// </summary>
        public void ClearTokenCache()
        {
            ClearTokenCacheAsync().GetAwaiter().GetResult();
        }

        /// <summary>
        /// Clears the internal in-memory token cache used by MSAL
        /// </summary>
        public async Task ClearTokenCacheAsync()
        {
            if (publicClientApplication != null)
            {
                var accounts = (await publicClientApplication.GetAccountsAsync().ConfigureAwait(false)).ToList();
                while (accounts.Any())
                {
                    await publicClientApplication.RemoveAsync(accounts.First()).ConfigureAwait(false);
                    accounts = (await publicClientApplication.GetAccountsAsync().ConfigureAwait(false)).ToList();
                }
            }
            if (confidentialClientApplication != null)
            {
#pragma warning disable CS0618 // Type or member is obsolete
                var accounts = (await confidentialClientApplication.GetAccountsAsync().ConfigureAwait(false)).ToList();
#pragma warning restore CS0618 // Type or member is obsolete
                while (accounts.Any())
                {
                    await confidentialClientApplication.RemoveAsync(accounts.First()).ConfigureAwait(false);
#pragma warning disable CS0618 // Type or member is obsolete
                    accounts = (await confidentialClientApplication.GetAccountsAsync().ConfigureAwait(false)).ToList();
#pragma warning restore CS0618 // Type or member is obsolete
                }
            }
        }

        /// <summary>
        /// called when disposing the object
        /// </summary>
        /// <param name="disposing"></param>
        protected virtual void Dispose(bool disposing)
        {
            // For backwards compatibility
        }

        /// <summary>
        /// Dispose the object
        /// </summary>
        public void Dispose()
        {
            // Do not change this code. Put cleanup code in 'Dispose(bool disposing)' method
            Dispose(disposing: true);
            GC.SuppressFinalize(this);
        }

        public PublicClientApplicationBuilder GetBuilderWithAuthority(PublicClientApplicationBuilder builder, AzureEnvironment azureEnvironment)
        {
            if (azureEnvironment == AzureEnvironment.Production)
            {
                var azureADEndPoint = GetAzureADLoginEndPoint(azureEnvironment);
                builder = builder.WithAuthority($"{azureADEndPoint}/organizations");
            }
            else
            {
                switch (azureEnvironment)
                {
                    case AzureEnvironment.USGovernment:
                    case AzureEnvironment.USGovernmentDoD:
                    case AzureEnvironment.USGovernmentHigh:
                        {
                            builder = builder.WithAuthority(AzureCloudInstance.AzureUsGovernment, AadAuthorityAudience.AzureAdMyOrg);
                            break;
                        }
                    case AzureEnvironment.Germany:
                        {
                            builder = builder.WithAuthority(AzureCloudInstance.AzureGermany, AadAuthorityAudience.AzureAdMyOrg);
                            break;
                        }
                    case AzureEnvironment.China:
                        {
                            builder = builder.WithAuthority(AzureCloudInstance.AzureChina, AadAuthorityAudience.AzureAdMyOrg);
                            break;
                        }
                }
            }
            return builder;
        }

        public ConfidentialClientApplicationBuilder GetBuilderWithAuthority(ConfidentialClientApplicationBuilder builder, AzureEnvironment azureEnvironment, string tenantId = "")
        {
            if (azureEnvironment == AzureEnvironment.Production)
            {
                var azureADEndPoint = GetAzureADLoginEndPoint(azureEnvironment);
                if (!string.IsNullOrEmpty(tenantId))
                {
                    builder = builder.WithAuthority($"{azureADEndPoint}/organizations", tenantId);
                }
                else
                {
                    builder = builder.WithAuthority($"{azureADEndPoint}/organizations");
                }
            }
            else
            {
                switch (azureEnvironment)
                {
                    case AzureEnvironment.USGovernment:
                    case AzureEnvironment.USGovernmentDoD:
                    case AzureEnvironment.USGovernmentHigh:
                        {
                            builder = builder.WithAuthority(AzureCloudInstance.AzureUsGovernment, AadAuthorityAudience.AzureAdMyOrg);
                            break;
                        }
                    case AzureEnvironment.Germany:
                        {
                            builder = builder.WithAuthority(AzureCloudInstance.AzureGermany, AadAuthorityAudience.AzureAdMyOrg);
                            break;
                        }
                    case AzureEnvironment.China:
                        {
                            builder = builder.WithAuthority(AzureCloudInstance.AzureChina, AadAuthorityAudience.AzureAdMyOrg);
                            break;
                        }
                }
            }
            return builder;
        }

        #endregion
    }
}
