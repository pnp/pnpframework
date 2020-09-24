namespace PnP.Framework.Utilities.Context
{
    internal enum ClientContextType
    {
        SharePointACSAppOnly = 0,
        AzureADCredentials = 1,
        AzureADCertificate = 2,
        Cookie = 3,
        AzureADInteractive = 4
    }
}
