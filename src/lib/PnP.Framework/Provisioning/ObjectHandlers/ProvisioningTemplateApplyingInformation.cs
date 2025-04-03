﻿using PnP.Framework.Provisioning.Model;
using PnP.Framework.Provisioning.Model.Configuration;
using System;
using System.Collections.Generic;

namespace PnP.Framework.Provisioning.ObjectHandlers
{
    /// <summary>
    /// Privisioning Progress Delegate
    /// </summary>
    /// <param name="message">Provisioning message</param>
    /// <param name="step"></param>
    /// <param name="total">Handlers count</param>
    public delegate void ProvisioningProgressDelegate(string message, int step, int total);

    /// <summary>
    /// Provisioning Messages Delegate
    /// </summary>
    /// <param name="message">Provisioning message</param>
    /// <param name="messageType">Provisioning message type</param>
    public delegate void ProvisioningMessagesDelegate(string message, ProvisioningMessageType messageType);

    /// <summary>
    /// Message to notify the completed provisioning of a site
    /// </summary>
    /// <param name="title">The Title of the provisioned site</param>
    /// <param name="url">The URL of the provisioned site</param>
    public delegate void ProvisioningSiteProvisionedDelegate(string title, string url);

    /// <summary>
    /// Handles methods for applying provisioning templates
    /// </summary>
    public partial class ProvisioningTemplateApplyingInformation
    {
        private Handlers handlersToProcess = Handlers.All;
        private List<ExtensibilityHandler> extensibilityHandlers = new List<ExtensibilityHandler>();
        private Dictionary<String, String> _accessTokens;

        public ProvisioningProgressDelegate ProgressDelegate { get; set; }

        public ProvisioningMessagesDelegate MessagesDelegate { get; set; }

        public ProvisioningSiteProvisionedDelegate SiteProvisionedDelegate { get; set; }

        public Dictionary<string, string> Parameters { get; set; }

        internal ApplyConfiguration ApplyConfiguration { get; set; }

        /// <summary>
        /// If true then persists template information
        /// </summary>
		public bool PersistTemplateInfo { get; set; } = true;

        /// <summary>
        /// If true, system propertybag entries that start with _, vti_, dlc_ etc. will be overwritten if overwrite = true on the PropertyBagEntry. If not true those keys will be skipped, regardless of the overwrite property of the entry.
        /// </summary>
        public bool OverwriteSystemPropertyBagValues { get; set; }

        /// <summary>
        /// If true, existing navigation nodes of the site (where applicable) will be cleared out before applying the navigation nodes from the template (if any). This setting will override any settings made in the template.
        /// </summary>
        public bool ClearNavigation { get; set; }

        /// <summary>
        /// If true then duplicate id errors when the same importing datarows simply generate a warning don't stop the engine. Reason for this is being able to apply the same template multiple times (Delta handling)
        /// without that failing cause the same record is being added twice
        /// </summary>
        public bool IgnoreDuplicateDataRowErrors { get; set; }

        /// <summary>
        /// If true then any content types in the template will be provisioned to subwebs
        /// </summary>
        public bool ProvisionContentTypesToSubWebs { get; set; }

        /// <summary>
        /// If true then any fields in the template will be provisioned to subwebs
        /// </summary>
        public bool ProvisionFieldsToSubWebs { get; set; }

        /// <summary>
        /// Specifies whether to also load site collection term groups when initializing the <see cref="TokenParser"/>. If
        /// <c>false</c>, only normal term groups will be loaded. This does not affect loading the site collection term group
        /// when one of the <c>sitecollectionterm</c> tokens was found.
        /// </summary>
        public bool LoadSiteCollectionTermGroups { get; set; } = true;

        /// <summary>
        /// Lists of Handlers to process
        /// </summary>
        public Handlers HandlersToProcess
        {
            get
            {
                return handlersToProcess;
            }
            set
            {
                handlersToProcess = value;
            }
        }

        /// <summary>
        /// List of ExtensibilityHandlers
        /// </summary>
        public List<ExtensibilityHandler> ExtensibilityHandlers
        {
            get
            {
                return extensibilityHandlers;
            }

            set
            {
                extensibilityHandlers = value;
            }
        }

        /// <summary>
        /// Allows to provide a dictionary of custom OAuth access tokens
        /// when working across different URLs during provisioning and
        /// using OAuth for AuthN/AuthZ
        /// </summary>
        public Dictionary<String, String> AccessTokens
        {
            get
            {
                if (this._accessTokens == null)
                {
                    this._accessTokens = new Dictionary<string, string>();
                }
                return (this._accessTokens);
            }
            set
            {
                this._accessTokens = value;
            }
        }

        /// <summary>
        /// Defines a delay to wait for after modern site creation
        /// </summary>
        public Int32 DelayAfterModernSiteCreation { get; set; }
    }
}
