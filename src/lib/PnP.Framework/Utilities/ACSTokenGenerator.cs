using SharePointPnP.IdentityModel.Extensions.S2S;
using SharePointPnP.IdentityModel.Extensions.S2S.Protocols.OAuth2;
using System;
using System.Threading.Tasks;

namespace PnP.Framework.Utilities
{
    internal sealed class ACSTokenGenerator
    {
        private readonly TokenHelper tokenHelper;
        private OAuth2AccessTokenResponse lastToken;

        /// <summary>
        /// Creates an OAuthAuthenticationProvider using ACS authentication.
        /// </summary>
        /// <param name="siteUrl">Site collection URL</param>
        /// <param name="options">Options object for the TokenHelper class</param>
        /// <returns>OAuthAuthenticationProvider that creates Tokens for ACS authentication</returns>
        private static async Task<ACSTokenGenerator> CreateAsync(Uri siteUrl, TokenHelperOptions options)
        {
            if (siteUrl == null) new ArgumentNullException(nameof(siteUrl));
            if (options == null) new ArgumentNullException(nameof(options));

            // realm is optional, determine it if not supplied
            if (string.IsNullOrEmpty(options.Realm))
            {
                options.Realm = await TokenHelper.GetRealmFromTargetUrlAsync(siteUrl);
            }

            return new ACSTokenGenerator(new TokenHelper(options));
        }

        private ACSTokenGenerator(TokenHelper tokenHelper)
        {
            this.tokenHelper = tokenHelper ?? throw new ArgumentNullException(nameof(tokenHelper));
        }

        public async Task<string> GetTokenAsync(Uri siteUrl)
        {
            // implement simple token caching
            if (lastToken == null || DateTime.Now.AddMinutes(3) >= lastToken.ExpiresOn)
            {
                lastToken = await tokenHelper.GetAppOnlyAccessTokenAsync(TokenHelper.SharePointPrincipal, siteUrl.Authority);
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
        public static async Task<ACSTokenGenerator> GetACSAuthenticationProviderAsync(Uri siteUrl, string realm, string appId, string appSecret, string acsHostUrl = "accesscontrol.windows.net", string globalEndPointPrefix = "accounts")
        {
            return await CreateAsync(siteUrl, new TokenHelperOptions()
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