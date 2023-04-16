using PnP.Framework;
using PnP.Framework.Utilities.Context;
using System;

namespace Microsoft.SharePoint.Client
{
    /// <summary>
    /// Class that holds the extension methods used to "tag" a client context for cloning support
    /// </summary>
    public static partial class InternalClientContextExtensions
    {
        private const string PnPSettingsKey = "SharePointPnP$Settings$ContextCloning";

        public static void AddContextSettings(this ClientRuntimeContext clientContext, ClientContextSettings contextData)
        {
            clientContext.StaticObjects[PnPSettingsKey] = contextData;
        }

        public static ClientContextSettings GetContextSettings(this ClientRuntimeContext clientContext)
        {
            if (!clientContext.StaticObjects.TryGetValue(PnPSettingsKey, out object settingsObject))
            {
                return null;
            }

            return (ClientContextSettings)settingsObject;
        }

        public static AzureEnvironment GetAzureEnvironment(this ClientRuntimeContext clientContext)
        {
            if (!clientContext.StaticObjects.TryGetValue(PnPSettingsKey, out object settingsObject))
            {
                // Do a best effort guess by determining the Environment based upon the url (if available)

                if (!string.IsNullOrEmpty(clientContext.Url))
                {
                    var url = new Uri(clientContext.Url);
                    var dnsDomain = url.Authority.Substring(url.Authority.LastIndexOf('.') + 1);

                    return dnsDomain switch
                    {
                        "com" => AzureEnvironment.Production,
                        "us" => AzureEnvironment.USGovernment,
                        "de" => AzureEnvironment.Germany,
                        "cn" => AzureEnvironment.China,
                        _ => AzureEnvironment.Production,
                    };
                }
            }
            else
            {
                return ((ClientContextSettings)settingsObject).Environment;
            }

            // If all fails, we assume production
            return AzureEnvironment.Production;
        }

    }
}
