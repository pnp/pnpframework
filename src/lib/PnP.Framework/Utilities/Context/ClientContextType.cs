namespace PnP.Framework.Utilities.Context
{
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
        PnPCoreSdk = 9
    }
}
