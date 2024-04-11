namespace PnP.Framework.Utilities.Context
{
    /// <summary>
    /// The authentication type used for setting up a ClientContext
    /// </summary>
    public enum ClientContextType
    {
        SharePointACSAppOnly = 0,
        AzureADCredentials = 1,
        AzureADCertificate = 2,
        Cookie = 3,
        AzureADInteractive = 4,
        AzureOnBehalfOf = 5,
        DeviceLogin = 6,
        OnPremises = 7,
        AccessToken = 8,
        PnPCoreSdk = 9,
        
        /// <summary>
        /// System Assigned Managed Identity in Azure
        /// </summary>
        SystemAssignedManagedIdentity = 10,

        /// <summary>
        /// User Assigned Managed Identity in Azure
        /// </summary>
        UserAssignedManagedIdentity = 11
    }
}
