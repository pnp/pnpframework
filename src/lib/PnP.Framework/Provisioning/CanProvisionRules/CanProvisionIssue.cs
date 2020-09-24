using System;

namespace PnP.Framework.Provisioning.CanProvisionRules
{
    /// <summary>
    /// Defines a CanProvision Issue item
    /// </summary>
    public class CanProvisionIssue
    {
        /// <summary>
        /// The Source of the CanProvision Issue
        /// </summary>
        public string Source { get; set; }

        /// <summary>
        /// Provides a text-based description of the Issue
        /// </summary>
        public string Message { get; set; }

        /// <summary>
        /// Provides a unique Tag for the current issue
        /// </summary>
        public CanProvisionIssueTags Tag { get; set; }

        /// <summary>
        /// Provides the Message of the source Exception of an Issue
        /// </summary>
        public string ExceptionMessage { get; set; }

        /// <summary>
        /// Provides the StackTrace of the source Exception of an Issue
        /// </summary>
        public string ExceptionStackTrace { get; set; }
    }
}
