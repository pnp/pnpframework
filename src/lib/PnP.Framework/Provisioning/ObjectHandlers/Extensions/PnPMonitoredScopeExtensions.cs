using PnP.Framework.Diagnostics;

namespace PnP.Framework.Provisioning.ObjectHandlers.Extensions
{
    internal static class PnPMonitoredScopeExtensions
    {
        public static void LogPropertyUpdate(this PnPMonitoredScope scope, string propertyName)
        {
            scope.LogDebug(CoreResources.PnPMonitoredScopeExtensions_LogPropertyUpdate_Updating_property__0_, propertyName);
        }
    }
}
