using PnP.Framework.Utilities.Context;

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
    }
}
