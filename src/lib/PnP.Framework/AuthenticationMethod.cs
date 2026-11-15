namespace PnP.Framework
{
    /// <summary>The authentication provider configured on an authentication manager, not the source of an individual token.</summary>
    public enum AuthenticationMethod
    {
        /// <summary>The provider is not known.</summary>
        Unknown,
        /// <summary>Interactive browser authentication.</summary>
        Browser,
        /// <summary>Operating system broker authentication.</summary>
        WAM,
        /// <summary>Device code authentication.</summary>
        DeviceCode,
        /// <summary>Username and password authentication.</summary>
        UsernamePassword,
        /// <summary>Certificate-based application authentication.</summary>
        Certificate,
        /// <summary>Legacy SharePoint ACS client secret authentication.</summary>
        ClientSecret,
        /// <summary>Azure managed identity authentication.</summary>
        ManagedIdentity,
        /// <summary>Workload identity authentication.</summary>
        WorkloadIdentity,
        /// <summary>Federated identity authentication.</summary>
        FederatedIdentity,
        /// <summary>A caller-supplied access token.</summary>
        AccessToken,
        /// <summary>On-premises credential authentication.</summary>
        OnPremises,
        /// <summary>Authentication supplied by another provider.</summary>
        ExternalProvider
    }
}
