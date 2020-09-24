using System;

namespace PnP.Framework.Provisioning.CanProvisionRules
{
    /// <summary>
    /// Basic interface for any CanProvision Rule
    /// </summary>
    interface ICanProvisionRuleBase
    {
        /// <summary>
        /// The name of the CanProvision Rule
        /// </summary>
        string Name { get; }
    }
}
