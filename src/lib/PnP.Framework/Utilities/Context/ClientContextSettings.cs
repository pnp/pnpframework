using System;
using System.Security.Cryptography.X509Certificates;

namespace PnP.Framework.Utilities.Context
{
    internal class ClientContextSettings
    {
        #region properties
        // Generic
        internal ClientContextType Type { get; set; }
        internal string SiteUrl { get; set; }
        internal AuthenticationManager AuthenticationManager { get; set; }
        #endregion

        #region methods
        internal bool UsesDifferentAudience(string newSiteUrl)
        {
            Uri newAudience = new Uri(newSiteUrl);
            Uri currentAudience = new Uri(this.SiteUrl);

            if (newAudience.Host != currentAudience.Host)
            {
                return true;
            }
            else
            {
                return false;
            }
        }
        #endregion


    }
}
