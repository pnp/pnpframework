using Microsoft.SharePoint.Client;
using System;
using System.Linq;
using System.Reflection;

namespace PnP.Framework.Utilities
{
    /// <summary>
    /// Holds PnP Core library identification tag and user-agent, and a tool to get tenant administration url based the URL of the web
    /// </summary>
    public static class PnPCoreUtilities
    {
        /// <summary>
        /// Gets a tag that identifies the PnP Core library
        /// </summary>
        /// <returns>PnP Core library identification tag</returns>
        public static string PnPCoreVersionTag
        {
            get
            {
                return (PnPCoreVersionTagLazy.Value);
            }
        }

        private static readonly Lazy<string> PnPCoreVersionTagLazy = new Lazy<string>(
            () =>
            {
                Assembly coreAssembly = Assembly.GetExecutingAssembly();
                var version = ((AssemblyFileVersionAttribute)coreAssembly.GetCustomAttribute(typeof(AssemblyFileVersionAttribute))).Version.Split('.');
                return $"PnPCore:{version[0]}.{version[1]}";
            },
            true);

        /// <summary>
        /// Gets a tag that identifies the PnP Core library for UserAgent string
        /// </summary>
        /// <returns>PnP Core library identification user-agent</returns>
        public static string PnPCoreUserAgent
        {
            get
            {
                return (PnPCoreUserAgentLazy.Value);
            }
        }

        private static readonly Lazy<string> PnPCoreUserAgentLazy = new Lazy<string>(
            () =>
            {
                Assembly coreAssembly = Assembly.GetExecutingAssembly();
                string result = $"NONISV|SharePointPnP|PnPCore/{((AssemblyFileVersionAttribute)coreAssembly.GetCustomAttribute(typeof(AssemblyFileVersionAttribute))).Version} ({Environment.OSVersion})";
                return (result);
            },
            true);

        /// <summary>
        /// Returns the tenant administration url based upon the URL of the web
        /// </summary>
        /// <param name="web"></param>
        /// <returns></returns>
        public static string GetTenantAdministrationUrl(this Web web)
        {
            var url = web.EnsureProperty(w => w.Url);

            return AuthenticationManager.GetTenantAdministrationUrl(url);
        }
    }
}
