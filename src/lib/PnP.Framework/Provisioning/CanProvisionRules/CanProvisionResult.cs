using System;
using System.Collections.Generic;

namespace PnP.Framework.Provisioning.CanProvisionRules
{
    /// <summary>
    /// Provides the complex output of the CanProvision method
    /// </summary>
    public class CanProvisionResult
    {
        /// <summary>
        /// Defines whether the Provisioning can proceed or not
        /// </summary>
        public Boolean CanProvision { get; set; } = true;

        /// <summary>
        /// The list of detailed CanProvision Issues, if any
        /// </summary>
        public List<CanProvisionIssue> Issues { get; set; } = new List<CanProvisionIssue>();
    }
}
