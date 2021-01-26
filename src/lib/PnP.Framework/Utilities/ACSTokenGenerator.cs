using SharePointPnP.IdentityModel.Extensions.S2S;
using SharePointPnP.IdentityModel.Extensions.S2S.Protocols.OAuth2;
using System;

namespace PnP.Framework.Utilities
{
    internal class ACSTokenGenerator
    {
        private TokenHelper tokenHelper;
        private OAuth2AccessTokenResponse lastToken;

        /// <summary>
        /// Creates an OAuthAuthenticationProvider using ACS authentication.
        /// </summary>
        /// <param name="siteUrl">Site collection URL</param>
        /// <param name="options">Options object for the TokenHelper class</param>
        /// <returns>OAuthAuthenticationProvider that creates Tokens for ACS authentication</returns>
        public ACSTokenGenerator(Uri siteUrl, TokenHelperOptions options)
        {
            if (siteUrl == null) new ArgumentNullException(nameof(siteUrl));
            if (options == null) new ArgumentNullException(nameof(options));

            // realm is optional, determine it if not supplied
            if (string.IsNullOrEmpty(options.Realm))
            {
                options.Realm = TokenHelper.GetRealmFromTargetUrl(siteUrl);
            }

            this.tokenHelper = new TokenHelper(options);
        }

        public string GetToken(Uri siteUrl)
        {
            // implement simple token caching
            if (lastToken == null || DateTime.Now.AddMinutes(3) >= lastToken.ExpiresOn)
            {
                lastToken = tokenHelper.GetAppOnlyAccessToken(TokenHelper.SharePointPrincipal, siteUrl.Authority);
            }

            return lastToken.AccessToken;
        }

        /// <summary>
        /// Creates an OAuthAuthenticationProvider usind ACS authentication.
        /// </summary>
        /// <param name="siteUrl">Site collection URL</param>
        /// <param name="realm">Tenant realm, may be null</param>
        /// <param name="appId">Application ID</param>
        /// <param name="appSecret">application Secret</param>
        /// <param name="acsHostUrl">ACS host url, may be null</param>
        /// <param name="globalEndPointPrefix">ACS endpoint prefix, may be null</param>
        /// <returns>OAuthAuthenticationProvider that creates Tokens for ACS authentication</returns>
        public static ACSTokenGenerator GetACSAuthenticationProvider(Uri siteUrl, string realm, string appId, string appSecret, string acsHostUrl = "accesscontrol.windows.net", string globalEndPointPrefix = "accounts")
        {
            return new ACSTokenGenerator(siteUrl, new TokenHelperOptions()
            {
                ClientId = appId,
                ClientSecret = appSecret,
                AcsHostUrl = acsHostUrl,
                GlobalEndPointPrefix = globalEndPointPrefix,
                Realm = realm
            });
        }
    }
}