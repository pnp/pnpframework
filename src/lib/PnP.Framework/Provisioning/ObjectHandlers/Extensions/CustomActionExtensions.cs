using PnP.Framework.Provisioning.Model;
using System;

namespace PnP.Framework.Provisioning.ObjectHandlers.Extensions
{
    internal static class CustomActionExtensions
    {
        /// <summary>
        /// Returns true when the given custom action is an SPFx extension.
        /// Every SPFx extension has a non-empty ClientSideComponentId.
        /// </summary>
        internal static bool IsSPFxCustomAction(this CustomAction customAction)
        {
            return customAction.ClientSideComponentId != Guid.Empty;
        }
    }
}
