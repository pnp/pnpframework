namespace PnP.Framework.TimerJobs.Enums
{
    /// <summary>
    /// Type of authentication, supports Office365, NetworkCredentials (on-premises) and AppOnly (both Office 365 as On-premises)
    /// </summary>
    public enum AuthenticationType
    {
        /// <summary>
        /// Office365 Authentication.
        /// </summary>
        Office365 = 0,
        /// <summary>
        /// Apps-Only Authentication.
        /// </summary>
        AppOnly = 2,
        /// <summary>
        /// Azure Active Directory Apps-Only Authentication.
        /// </summary>
        AzureADAppOnly = 3,
        /// <summary>
        /// Consumer provides a valid access token
        /// </summary>
        AccessToken = 4,
    }
}
