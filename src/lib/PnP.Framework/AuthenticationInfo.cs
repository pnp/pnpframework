namespace PnP.Framework
{
    /// <summary>
    /// An immutable, credential-free snapshot of a provider and its latest successful acquisition. When a
    /// manager serves multiple resources or connections, identity fields describe its latest acquisition,
    /// not an identity pinned to each resource. It does not describe resource permissions,
    /// token expiry, or whether another token can be acquired without interaction.
    /// </summary>
    public sealed class AuthenticationInfo
    {
        /// <summary>Creates a snapshot. Unavailable identity information must be omitted.</summary>
        /// <param name="method">The configured provider.</param>
        /// <param name="identityType">The identity type, when known.</param>
        /// <param name="accountId">The MSAL home-account identifier, when available.</param>
        /// <param name="username">The display username returned by MSAL.</param>
        /// <param name="homeTenantId">The home-account tenant returned by MSAL.</param>
        /// <param name="authenticationTenantId">The resolved tenant returned by successful token acquisition.</param>
        public AuthenticationInfo(AuthenticationMethod method = AuthenticationMethod.Unknown,
            AuthenticationIdentityType identityType = AuthenticationIdentityType.Unknown,
            string accountId = null, string username = null, string homeTenantId = null,
            string authenticationTenantId = null)
        {
            Method = method;
            IdentityType = identityType;
            AccountId = accountId;
            Username = username;
            HomeTenantId = homeTenantId;
            AuthenticationTenantId = authenticationTenantId;
        }

        /// <summary>The configured provider, including when its tokens are acquired silently.</summary>
        public AuthenticationMethod Method { get; }
        /// <summary>The identity type; a supplied token need not disclose its identity.</summary>
        public AuthenticationIdentityType IdentityType { get; }
        /// <summary>The MSAL home-account identifier, not a tenant-local user object ID.</summary>
        public string AccountId { get; }
        /// <summary>A display username, not an immutable identity or authorization key.</summary>
        public string Username { get; }
        /// <summary>The user's home tenant, which can differ from the authentication tenant for guests.</summary>
        public string HomeTenantId { get; }
        /// <summary>The resolved tenant from successful MSAL acquisition, not an unverified input tenant name.</summary>
        public string AuthenticationTenantId { get; }
    }
}
