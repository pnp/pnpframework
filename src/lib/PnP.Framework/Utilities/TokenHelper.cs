using Microsoft.IdentityModel.Tokens;
using Microsoft.SharePoint.Client;
using Newtonsoft.Json;
using SharePointPnP.IdentityModel.Extensions.S2S.Protocols.OAuth2;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Globalization;
using System.IdentityModel.Tokens.Jwt;
using System.IO;
using System.Linq;
using System.Net;
using System.Security.Claims;
using System.Security.Cryptography.X509Certificates;
using System.Text;

namespace PnP.Framework.Utilities
{
    internal static class TokenHelper
    {
        #region public fields

        /// <summary>
        /// SharePoint principal.
        /// </summary>
        public const string SharePointPrincipal = "00000003-0000-0ff1-ce00-000000000000";

        /// <summary>
        /// Lifetime of HighTrust access token, 12 hours.
        /// </summary>
        public static readonly TimeSpan HighTrustAccessTokenLifetime = TimeSpan.FromHours(12.0);

        #endregion public fields

        #region private fields

        //
        // Configuration Constants
        //

        private const string AuthorizationPage = "_layouts/15/OAuthAuthorize.aspx";
        private const string RedirectPage = "_layouts/15/AppRedirect.aspx";
        private const string AcsPrincipalName = "00000001-0000-0000-c000-000000000000";
        private const string AcsMetadataEndPointRelativeUrl = "metadata/json/1";
        private const string S2SProtocol = "OAuth2";
        private const string DelegationIssuance = "DelegationIssuance1.0";
        private const string NameIdentifierClaimType = "nameid";
        private const string TrustedForImpersonationClaimType = "trustedfordelegation";
        private const string ActorTokenClaimType = "actortoken";

        //
        // Hosted app configuration
        //
        private static string clientId = null;
        private static string issuerId = null;
        private static string hostedAppHostNameOverride = null;
        private static string hostedAppHostName = null;
        private static string clientSecret = null;
        private static string secondaryClientSecret = null;
        private static string realm = null;
        private static string serviceNamespace = null;
        private static string identityClaimType = null;
        private static string trustedIdentityTokenIssuerName = null;
        //
        // Environment Constants
        //
        private static string acsHostUrl = "accesscontrol.windows.net";
        private static string globalEndPointPrefix = "accounts";

        public static string AcsHostUrl
        {
            get
            {
                if (String.IsNullOrEmpty(acsHostUrl))
                {
                    return "accesscontrol.windows.net";
                }
                else
                {
                    return acsHostUrl;
                }
            }
            set
            {
                acsHostUrl = value;
            }
        }

        public static string GlobalEndPointPrefix
        {
            get
            {
                if (globalEndPointPrefix == null)
                {
                    return "accounts";
                }
                else
                {
                    return globalEndPointPrefix;
                }
            }
            set
            {
                globalEndPointPrefix = value;
            }
        }

        public static string ClientId
        {
            get
            {
                if (String.IsNullOrEmpty(clientId))
                {
                    return string.IsNullOrEmpty(ConfigurationManager.AppSettings.Get("ClientId")) ? ConfigurationManager.AppSettings.Get("HostedAppName") : ConfigurationManager.AppSettings.Get("ClientId");
                }
                else
                {
                    return clientId;
                }
            }
            set
            {
                clientId = value;
            }
        }

        public static string IssuerId
        {
            get
            {
                if (String.IsNullOrEmpty(issuerId))
                {
                    return string.IsNullOrEmpty(ConfigurationManager.AppSettings.Get("IssuerId")) ? ClientId : ConfigurationManager.AppSettings.Get("IssuerId");
                }
                else
                {
                    return issuerId;
                }
            }
            set
            {
                issuerId = value;
            }
        }

        public static string HostedAppHostNameOverride
        {
            get
            {
                if (String.IsNullOrEmpty(hostedAppHostNameOverride))
                {
                    return ConfigurationManager.AppSettings.Get("HostedAppHostNameOverride");
                }
                else
                {
                    return hostedAppHostNameOverride;
                }
            }
            set
            {
                hostedAppHostNameOverride = value;
            }
        }

        public static string HostedAppHostName
        {
            get
            {
                if (String.IsNullOrEmpty(hostedAppHostName))
                {
                    return ConfigurationManager.AppSettings.Get("HostedAppHostName");
                }
                else
                {
                    return hostedAppHostName;
                }
            }
            set
            {
                hostedAppHostName = value;
            }
        }

        public static string ClientSecret
        {
            get
            {
                if (String.IsNullOrEmpty(clientSecret))
                {
                    return string.IsNullOrEmpty(ConfigurationManager.AppSettings.Get("ClientSecret")) ? ConfigurationManager.AppSettings.Get("HostedAppSigningKey") : ConfigurationManager.AppSettings.Get("ClientSecret");
                }
                else
                {
                    return clientSecret;
                }
            }
            set
            {
                clientSecret = value;
            }
        }

        public static string SecondaryClientSecret
        {
            get
            {
                if (String.IsNullOrEmpty(secondaryClientSecret))
                {
                    return ConfigurationManager.AppSettings.Get("SecondaryClientSecret");
                }
                else
                {
                    return secondaryClientSecret;
                }
            }
            set
            {
                secondaryClientSecret = value;
            }
        }

        public static string Realm
        {
            get
            {
                if (String.IsNullOrEmpty(realm))
                {
                    return ConfigurationManager.AppSettings.Get("Realm");
                }
                else
                {
                    return realm;
                }
            }
            set
            {
                realm = value;
            }
        }

        public static string ServiceNamespace
        {
            get
            {
                if (String.IsNullOrEmpty(serviceNamespace))
                {
                    return ConfigurationManager.AppSettings.Get("Realm");
                }
                else
                {
                    return serviceNamespace;
                }
            }
            set
            {
                serviceNamespace = value;
            }
        }

        public static string IdentityClaimType
        {
            get
            {
                if (String.IsNullOrEmpty(identityClaimType))
                {
                    return ConfigurationManager.AppSettings.Get("IdentityClaimType");
                }
                else
                {
                    return identityClaimType;
                }
            }
            set
            {
                identityClaimType = value;
            }
        }

        public static string TrustedIdentityTokenIssuerName
        {
            get
            {
                if (String.IsNullOrEmpty(trustedIdentityTokenIssuerName))
                {
                    return ConfigurationManager.AppSettings.Get("TrustedIdentityTokenIssuerName");
                }
                else
                {
                    return trustedIdentityTokenIssuerName;
                }
            }
            set
            {
                trustedIdentityTokenIssuerName = value;
            }
        }

        //private static readonly string ClientId = string.IsNullOrEmpty(WebConfigurationManager.AppSettings.Get("ClientId")) ? WebConfigurationManager.AppSettings.Get("HostedAppName") : WebConfigurationManager.AppSettings.Get("ClientId");
        //private static readonly string IssuerId = string.IsNullOrEmpty(WebConfigurationManager.AppSettings.Get("IssuerId")) ? ClientId : WebConfigurationManager.AppSettings.Get("IssuerId");
        //private static readonly string HostedAppHostNameOverride = WebConfigurationManager.AppSettings.Get("HostedAppHostNameOverride");
        //private static readonly string HostedAppHostName = WebConfigurationManager.AppSettings.Get("HostedAppHostName");
        //private static readonly string ClientSecret = string.IsNullOrEmpty(WebConfigurationManager.AppSettings.Get("ClientSecret")) ? WebConfigurationManager.AppSettings.Get("HostedAppSigningKey") : WebConfigurationManager.AppSettings.Get("ClientSecret");
        //private static readonly string SecondaryClientSecret = WebConfigurationManager.AppSettings.Get("SecondaryClientSecret");
        //private static readonly string Realm = WebConfigurationManager.AppSettings.Get("Realm");
        //private static readonly string ServiceNamespace = WebConfigurationManager.AppSettings.Get("Realm");

        private static readonly string ClientSigningCertificatePath = ConfigurationManager.AppSettings.Get("ClientSigningCertificatePath");
        private static readonly string ClientSigningCertificatePassword = ConfigurationManager.AppSettings.Get("ClientSigningCertificatePassword");
        public static X509Certificate2 ClientCertificate = (string.IsNullOrEmpty(ClientSigningCertificatePath) || string.IsNullOrEmpty(ClientSigningCertificatePassword)) ? null : new X509Certificate2(ClientSigningCertificatePath, ClientSigningCertificatePassword);
        private static SigningCredentials SigningCredentials
        {
            get
            {
                var securityKey = new Microsoft.IdentityModel.Tokens.SymmetricSecurityKey(ClientCertificate.GetPublicKey());
                return (ClientCertificate == null) ? null : new SigningCredentials(securityKey, SecurityAlgorithms.RsaSha256Signature, SecurityAlgorithms.Sha256Digest);
            }
        }

        #endregion

        public static string GetRealmFromTargetUrl(Uri targetApplicationUri)
        {
            WebRequest request = WebRequest.Create(targetApplicationUri.ToString().TrimEnd(new[] { '/' }) + "/_vti_bin/client.svc");
            request.Headers.Add("Authorization: Bearer ");

            try
            {
                using (request.GetResponse())
                {
                }
            }
            catch (WebException e)
            {
                if (e.Response == null)
                {
                    return null;
                }

                string bearerResponseHeader = e.Response.Headers["WWW-Authenticate"];
                if (string.IsNullOrEmpty(bearerResponseHeader))
                {
                    return null;
                }

                const string bearer = "Bearer realm=\"";
                int bearerIndex = bearerResponseHeader.IndexOf(bearer, StringComparison.Ordinal);
                if (bearerIndex < 0)
                {
                    return null;
                }

                int realmIndex = bearerIndex + bearer.Length;

                if (bearerResponseHeader.Length >= realmIndex + 36)
                {
                    string targetRealm = bearerResponseHeader.Substring(realmIndex, 36);

                    Guid realmGuid;

                    if (Guid.TryParse(targetRealm, out realmGuid))
                    {
                        return targetRealm;
                    }
                }
            }
            return null;
        }
        public static SharePointContextToken ReadAndValidateContextToken(string contextTokenString, string appHostName = null)
        {
            JwtSecurityTokenHandler tokenHandler = CreateJsonWebSecurityTokenHandler();
            SecurityToken securityToken = tokenHandler.ReadToken(contextTokenString);
            JwtSecurityToken jsonToken = securityToken as JwtSecurityToken;
            SharePointContextToken token = SharePointContextToken.Create(jsonToken);

            string stsAuthority = (new Uri(token.SecurityTokenServiceUri)).Authority;
            int firstDot = stsAuthority.IndexOf('.');

            GlobalEndPointPrefix = stsAuthority.Substring(0, firstDot);
            AcsHostUrl = stsAuthority.Substring(firstDot + 1);

            //SecurityToken validatedToken = null;

            //tokenHandler.ValidateToken(jsonToken.ToString(),TokenValidationParameters.DefaultAuthenticationType.ToString(),out validatedToken);

            string[] acceptableAudiences;
            if (!String.IsNullOrEmpty(HostedAppHostNameOverride))
            {
                acceptableAudiences = HostedAppHostNameOverride.Split(';');
            }
            else if (appHostName == null)
            {
                acceptableAudiences = new[] { HostedAppHostName };
            }
            else
            {
                acceptableAudiences = new[] { appHostName };
            }

            bool validationSuccessful = false;
            string realm = Realm ?? token.Realm;
            foreach (var audience in acceptableAudiences)
            {
                string principal = GetFormattedPrincipal(ClientId, audience, realm);
                if (StringComparer.OrdinalIgnoreCase.Equals(token.Audiences, principal))
                {
                    validationSuccessful = true;
                    break;
                }
            }

            if (!validationSuccessful)
            {
                throw new Exception(
                    String.Format(CultureInfo.CurrentCulture,
                    "\"{0}\" is not the intended audience \"{1}\"", String.Join(";", acceptableAudiences), token.Audiences));
            }

            return token;
        }
        /// <summary>
        /// Uses the specified access token to create a client context
        /// </summary>
        /// <param name="targetUrl">Url of the target SharePoint site</param>
        /// <param name="accessToken">Access token to be used when calling the specified targetUrl</param>
        /// <returns>A ClientContext ready to call targetUrl with the specified access token</returns>
        public static ClientContext GetClientContextWithAccessToken(string targetUrl, string accessToken)
        {
            ClientContext clientContext = new ClientContext(targetUrl);

            clientContext.ExecutingWebRequest +=
                delegate (object oSender, WebRequestEventArgs webRequestEventArgs)
                {
                    webRequestEventArgs.WebRequestExecutor.RequestHeaders["Authorization"] =
                        "Bearer " + accessToken;
                };

            return clientContext;
        }

        /// <summary>
        /// Retrieves an app-only access token from ACS to call the specified principal
        /// at the specified targetHost. The targetHost must be registered for target principal.  If specified realm is
        /// null, the "Realm" setting in web.config will be used instead.
        /// </summary>
        /// <param name="targetPrincipalName">Name of the target principal to retrieve an access token for</param>
        /// <param name="targetHost">Url authority of the target principal</param>
        /// <param name="targetRealm">Realm to use for the access token's nameid and audience</param>
        /// <returns>An access token with an audience of the target principal</returns>
        public static OAuth2AccessTokenResponse GetAppOnlyAccessToken(
            string targetPrincipalName,
            string targetHost,
            string targetRealm)
        {

            if (targetRealm == null)
            {
                targetRealm = Realm;
            }

            string resource = GetFormattedPrincipal(targetPrincipalName, targetHost, targetRealm);
            string clientId = GetFormattedPrincipal(ClientId, HostedAppHostName, targetRealm);

            OAuth2AccessTokenRequest oauth2Request = OAuth2MessageFactory.CreateAccessTokenRequestWithClientCredentials(clientId, ClientSecret, resource);
            oauth2Request.Resource = resource;

            // Get token
            OAuth2S2SClient client = new OAuth2S2SClient();

            OAuth2AccessTokenResponse oauth2Response;
            try
            {
                oauth2Response =
                    client.Issue(AcsMetadataParser.GetStsUrl(targetRealm), oauth2Request) as OAuth2AccessTokenResponse;
            }
            catch (WebException wex) when (wex.Response != null)
            {
                using (StreamReader sr = new StreamReader(wex.Response.GetResponseStream()))
                {
                    string responseText = sr.ReadToEnd();
                    throw new WebException(wex.Message + " - " + responseText, wex);
                }
            }

            return oauth2Response;
        }

        #region AcsMetadataParser

        // This class is used to get MetaData document from the global STS endpoint. It contains
        // methods to parse the MetaData document and get endpoints and STS certificate.
        public static class AcsMetadataParser
        {
            public static X509Certificate2 GetAcsSigningCert(string realm)
            {
                JsonMetadataDocument document = GetMetadataDocument(realm);

                if (null != document.keys && document.keys.Count > 0)
                {
                    JsonKey signingKey = document.keys[0];

                    if (null != signingKey && null != signingKey.keyValue)
                    {
                        return new X509Certificate2(Encoding.UTF8.GetBytes(signingKey.keyValue.value));
                    }
                }

                throw new Exception("Metadata document does not contain ACS signing certificate.");
            }

            public static string GetDelegationServiceUrl(string realm)
            {
                JsonMetadataDocument document = GetMetadataDocument(realm);

                JsonEndpoint delegationEndpoint = document.endpoints.SingleOrDefault(e => e.protocol == DelegationIssuance);

                if (null != delegationEndpoint)
                {
                    return delegationEndpoint.location;
                }
                throw new Exception("Metadata document does not contain Delegation Service endpoint Url");
            }

            private static JsonMetadataDocument GetMetadataDocument(string realm)
            {
                string acsMetadataEndpointUrlWithRealm = String.Format(CultureInfo.InvariantCulture, "{0}?realm={1}",
                                                                       GetAcsMetadataEndpointUrl(),
                                                                       realm);
                byte[] acsMetadata;
                using (WebClient webClient = new WebClient())
                {

                    acsMetadata = webClient.DownloadData(acsMetadataEndpointUrlWithRealm);
                }
                string jsonResponseString = Encoding.UTF8.GetString(acsMetadata);

                JsonMetadataDocument document = JsonConvert.DeserializeObject<JsonMetadataDocument>(jsonResponseString);

                if (null == document)
                {
                    throw new Exception("No metadata document found at the global endpoint " + acsMetadataEndpointUrlWithRealm);
                }

                return document;
            }

            public static string GetStsUrl(string realm)
            {
                JsonMetadataDocument document = GetMetadataDocument(realm);

                JsonEndpoint s2sEndpoint = document.endpoints.SingleOrDefault(e => e.protocol == S2SProtocol);

                if (null != s2sEndpoint)
                {
                    return s2sEndpoint.location;
                }

                throw new Exception("Metadata document does not contain STS endpoint URL");
            }

            private class JsonMetadataDocument
            {
                public string serviceName { get; set; }
                public List<JsonEndpoint> endpoints { get; set; }
                public List<JsonKey> keys { get; set; }
            }

            private class JsonEndpoint
            {
                public string location { get; set; }
                public string protocol { get; set; }
                public string usage { get; set; }
            }

            private class JsonKeyValue
            {
                public string type { get; set; }
                public string value { get; set; }
            }

            private class JsonKey
            {
                public string usage { get; set; }
                public JsonKeyValue keyValue { get; set; }
            }
        }

        #endregion


        #region private methods

        //private static ClientContext CreateAcsClientContextForUrl(SPRemoteEventProperties properties, Uri sharepointUrl)
        //{
        //    string contextTokenString = properties.ContextToken;

        //    if (String.IsNullOrEmpty(contextTokenString))
        //    {
        //        return null;
        //    }

        //    SharePointContextToken contextToken = ReadAndValidateContextToken(contextTokenString, OperationContext.Current.IncomingMessageHeaders.To.Host);
        //    string accessToken = GetAccessToken(contextToken, sharepointUrl.Authority).AccessToken;

        //    return GetClientContextWithAccessToken(sharepointUrl.ToString(), accessToken);
        //}

        private static string GetAcsMetadataEndpointUrl()
        {
            return Path.Combine(GetAcsGlobalEndpointUrl(), AcsMetadataEndPointRelativeUrl);
        }

        private static string GetFormattedPrincipal(string principalName, string hostName, string realm)
        {
            if (!String.IsNullOrEmpty(hostName))
            {
                return String.Format(CultureInfo.InvariantCulture, "{0}/{1}@{2}", principalName, hostName, realm);
            }

            return String.Format(CultureInfo.InvariantCulture, "{0}@{1}", principalName, realm);
        }

        private static string GetAcsPrincipalName(string realm)
        {
            return GetFormattedPrincipal(AcsPrincipalName, new Uri(GetAcsGlobalEndpointUrl()).Host, realm);
        }

        private static string GetAcsGlobalEndpointUrl()
        {
            if (GlobalEndPointPrefix.Length == 0)
            {
                return String.Format(CultureInfo.InvariantCulture, "https://{0}/", AcsHostUrl);
            }
            else
            {
                return String.Format(CultureInfo.InvariantCulture, "https://{0}.{1}/", GlobalEndPointPrefix, AcsHostUrl);
            }
        }

        public static JwtSecurityTokenHandler CreateJsonWebSecurityTokenHandler()
        {
            JwtSecurityTokenHandler handler = new JwtSecurityTokenHandler();

            //handler.Configuration = new SecurityTokenHandlerConfiguration();
            //handler.Configuration.AudienceRestriction = new AudienceRestriction(AudienceUriMode.Never);
            //handler.Configuration.CertificateValidator = X509CertificateValidator.None;

            //List<byte[]> securityKeys = new List<byte[]>();
            //securityKeys.Add(Convert.FromBase64String(ClientSecret));
            //if (!string.IsNullOrEmpty(SecondaryClientSecret))
            //{
            //    securityKeys.Add(Convert.FromBase64String(SecondaryClientSecret));
            //}


            //List<SecurityToken> securityTokens = new List<SecurityToken>();
            //securityTokens.Add(new MultipleSymmetricKeySecurityToken(securityKeys));

            //handler.Configuration.IssuerTokenResolver =
            //    SecurityTokenResolver.CreateDefaultSecurityTokenResolver(
            //    new ReadOnlyCollection<SecurityToken>(securityTokens),
            //    false);
            //SymmetricKeyIssuerNameRegistry issuerNameRegistry = new SymmetricKeyIssuerNameRegistry();
            //foreach (byte[] securitykey in securityKeys)
            //{
            //    issuerNameRegistry.AddTrustedIssuer(securitykey, GetAcsPrincipalName(ServiceNamespace));
            //}
            //handler.Configuration.IssuerNameRegistry = issuerNameRegistry;
            return handler;
        }

        private static string GetS2SAccessTokenWithClaims(
            string targetApplicationHostName,
            string targetRealm,
            IEnumerable<Claim> claims)
        {
            return IssueToken(
                ClientId,
                IssuerId,
                targetRealm,
                SharePointPrincipal,
                targetRealm,
                targetApplicationHostName,
                true,
                claims,
                claims == null);
        }

        //private static Claim[] GetClaimsWithWindowsIdentity(WindowsIdentity identity)
        //{
        //    JsonWebTokenClaim[] claims = new JsonWebTokenClaim[]
        //    {
        //        new JsonWebTokenClaim(NameIdentifierClaimType, identity.User.Value.ToLower()),
        //        new JsonWebTokenClaim("nii", "urn:office:idp:activedirectory")
        //    };
        //    return claims;
        //}

        //private static JsonWebTokenClaim[] GetClaimsWithClaimsIdentity(System.Security.Claims.ClaimsIdentity identity, string identityClaimType, string trustedProviderName)
        //{
        //	var identityClaim = identity.Claims.Where(c => string.Equals(c.Type, identityClaimType, StringComparison.InvariantCultureIgnoreCase)).First();
        //	JsonWebTokenClaim[] claims = new JsonWebTokenClaim[]
        //	{
        //		new JsonWebTokenClaim(NameIdentifierClaimType, identityClaim.Value),
        //		new JsonWebTokenClaim("nii", "trusted:" + trustedProviderName)
        //	};
        //	return claims;
        //}

        private static string IssueToken(
            string sourceApplication,
            string issuerApplication,
            string sourceRealm,
            string targetApplication,
            string targetRealm,
            string targetApplicationHostName,
            bool trustedForDelegation,
            IEnumerable<Claim> claims,
            bool appOnly = false)
        {

            #region Actor token

            string issuer = string.IsNullOrEmpty(sourceRealm) ? issuerApplication : string.Format("{0}@{1}", issuerApplication, sourceRealm);
            string nameid = string.IsNullOrEmpty(sourceRealm) ? sourceApplication : string.Format("{0}@{1}", sourceApplication, sourceRealm);
            string audience = string.Format("{0}/{1}@{2}", targetApplication, targetApplicationHostName, targetRealm);

            List<Claim> actorClaims = new List<Claim>
            {
                new Claim("nameid", nameid)
            };
            if (trustedForDelegation && !appOnly)
            {
                actorClaims.Add(new Claim(TrustedForImpersonationClaimType, "true"));
            }

            // Create token
            JwtSecurityToken actorToken = new JwtSecurityToken(
                issuer: issuer,
                audience: audience,
                notBefore: DateTime.UtcNow,
                expires: DateTime.UtcNow.Add(HighTrustAccessTokenLifetime),
                signingCredentials: SigningCredentials,
                claims: actorClaims);

            string actorTokenString = new JwtSecurityTokenHandler().WriteToken(actorToken);

            if (appOnly)
            {
                // App-only token is the same as actor token for delegated case
                return actorTokenString;
            }

            #endregion Actor token

            #region Outer token

            List<Claim> outerClaims = null == claims ? new List<Claim>() : new List<Claim>(claims);
            outerClaims.Add(new Claim(ActorTokenClaimType, actorTokenString));

            JwtSecurityToken jsonToken = new JwtSecurityToken(
                nameid, // outer token issuer should match actor token nameid
                audience,
                outerClaims,
                DateTime.UtcNow,
                DateTime.UtcNow.Add(HighTrustAccessTokenLifetime));

            string accessToken = new JwtSecurityTokenHandler().WriteToken(jsonToken);

            #endregion Outer token

            return accessToken;
        }

        #endregion
    }

    /// <summary>
    /// A JsonWebSecurityToken generated by SharePoint to authenticate to a 3rd party application and allow callbacks using a refresh token
    /// </summary>
    public class SharePointContextToken : JwtSecurityToken
    {
        /// <summary>
        /// Creates SharePoint ContextToken
        /// </summary>
        /// <param name="contextToken">JsonWebSecurityToken object</param>
        /// <returns>Returns SharePoint ContextToken object</returns>
        public static SharePointContextToken Create(JwtSecurityToken contextToken)
        {
            return new SharePointContextToken(contextToken.Issuer, contextToken.Audiences.FirstOrDefault(), contextToken.ValidFrom, contextToken.ValidTo, contextToken.Claims);
        }

        /// <summary>
        /// Constructor for SharePointContextToken class
        /// </summary>
        /// <param name="issuer">Token Issuer</param>
        /// <param name="audience">Token Audience</param>
        /// <param name="validFrom">Validity start date for token</param>
        /// <param name="validTo">Validity end date for token</param>
        /// <param name="claims">IEnumerable of JsonWebTokenClaim object</param>
        public SharePointContextToken(string issuer, string audience, DateTime validFrom, DateTime validTo, IEnumerable<Claim> claims)
            : base(issuer, audience, claims, validFrom, validTo)
        {
        }

        /*
        /// <summary>
        /// Constructor for SharePointContextToken class
        /// </summary>
        /// <param name="issuer">Token Issuer</param>
        /// <param name="audience">Token Audience</param>
        /// <param name="validFrom">Validity start date for token</param>
        /// <param name="validTo">Validity end date for token</param>
        /// <param name="claims">IEnumerable of JsonWebTokenClaim object</param>
        /// <param name="issuerToken">SecurityToken object</param>
        /// <param name="actorToken">JsonWebSecurityToken object</param>
        public SharePointContextToken(string issuer, string audience, DateTime validFrom, DateTime validTo, IEnumerable<Claim> claims, SecurityToken issuerToken, JwtSecurityToken actorToken)
            : base(issuer, audience, claims, validFrom, validTo, issuerToken, actorToken)
        {
        }
        */

        /// <summary>
        /// Constructor for SharePointContextToken class
        /// </summary>
        /// <param name="issuer">Token Issuer</param>
        /// <param name="audience">Token Audience</param>
        /// <param name="validFrom">Validity start date for token</param>
        /// <param name="validTo">Validity end date for token</param>
        /// <param name="claims">IEnumerable of JsonWebTokenClaim object</param>
        /// <param name="signingCredentials">SigningCredentials object</param>
        public SharePointContextToken(string issuer, string audience, DateTime validFrom, DateTime validTo, IEnumerable<Claim> claims, SigningCredentials signingCredentials)
            : base(issuer, audience, claims, validFrom, validTo, signingCredentials)
        {
        }

        /// <summary>
        /// The context token's "nameid" claim
        /// </summary>
        public string NameId
        {
            get
            {
                return GetClaimValue(this, "nameid");
            }
        }

        /// <summary>
        /// The principal name portion of the context token's "appctxsender" claim
        /// </summary>
        public string TargetPrincipalName
        {
            get
            {
                string appctxsender = GetClaimValue(this, "appctxsender");

                if (appctxsender == null)
                {
                    return null;
                }

                return appctxsender.Split('@')[0];
            }
        }

        /// <summary>
        /// The context token's "refreshtoken" claim
        /// </summary>
        public string RefreshToken
        {
            get
            {
                return GetClaimValue(this, "refreshtoken");
            }
        }

        /// <summary>
        /// The context token's "CacheKey" claim
        /// </summary>
        public string CacheKey
        {
            get
            {
                string appctx = GetClaimValue(this, "appctx");
                if (appctx == null)
                {
                    return null;
                }

                using (ClientContext ctx = new ClientContext("http://tempuri.org"))
                {
                    Dictionary<string, object> dict = (Dictionary<string, object>)ctx.ParseObjectFromJsonString(appctx);
                    string cacheKey = (string)dict["CacheKey"];
                    return cacheKey;
                }
            }
        }

        /// <summary>
        /// The context token's "SecurityTokenServiceUri" claim
        /// </summary>
        public string SecurityTokenServiceUri
        {
            get
            {
                string appctx = GetClaimValue(this, "appctx");
                if (appctx == null)
                {
                    return null;
                }

                ClientContext ctx = new ClientContext("http://tempuri.org");
                Dictionary<string, object> dict = (Dictionary<string, object>)ctx.ParseObjectFromJsonString(appctx);
                string securityTokenServiceUri = (string)dict["SecurityTokenServiceUri"];

                return securityTokenServiceUri;
            }
        }

        /// <summary>
        /// The realm portion of the context token's "audience" claim
        /// </summary>
        public string Realm
        {
            get
            {
                string aud = Audiences.FirstOrDefault();
                if (aud == null)
                {
                    return null;
                }

                string tokenRealm = aud.Substring(aud.IndexOf('@') + 1);

                return tokenRealm;
            }
        }

        private static string GetClaimValue(JwtSecurityToken token, string claimType)
        {
            if (token == null)
            {
                throw new ArgumentNullException(nameof(token));
            }

            foreach (Claim claim in token.Claims)
            {
                if (StringComparer.Ordinal.Equals(claim.Type, claimType))
                {
                    return claim.Value;
                }
            }

            return null;
        }

    }
}
