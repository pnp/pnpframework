using System;

namespace Microsoft.SharePoint.Client
{

    /// <summary>
    /// Class that deals with branding features
    /// </summary>
    public static partial class BrandingExtensions
    {

        #region TO BE DEPRECATED IN MARCH 2016 RELEASE - Long deprecation time to avoid issues

        /// <summary>
        /// Disables the Responsive UI on a Classic SharePoint Site
        /// </summary>
        /// <param name="site">The Site to disable the Responsive UI on</param>
        [Obsolete("Use DisableResponsiveUI(site)")]
        public static void DisableReponsiveUI(this Site site)
        {
            try
            {
                site.DeleteJsLink("PnPResponsiveUI");
            }
            catch
            {
                // Swallow exception as responsive UI might not be active.
            }
        }

        /// <summary>
        /// Disables the Responsive UI on a Classic SharePoint Web
        /// </summary>
        /// <param name="web">The Web to disable the Responsive UI on</param>
        [Obsolete("Use DisableResponsiveUI(web)")]
        public static void DisableReponsiveUI(this Web web)
        {
            try
            {
                web.DeleteJsLink("PnPResponsiveUI");
            }
            catch
            {
                // Swallow exception as responsive UI might not be active.
            }
        }

        #endregion
    }
}