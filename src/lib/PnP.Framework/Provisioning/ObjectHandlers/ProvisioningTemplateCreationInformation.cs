﻿using Microsoft.SharePoint.Client;
using PnP.Framework.Provisioning.Connectors;
using PnP.Framework.Provisioning.Model;
using System;
using System.Collections.Generic;

namespace PnP.Framework.Provisioning.ObjectHandlers
{
    /// <summary>
    /// Handles methods for Provisioning Template Creation Information
    /// </summary>
    public class ProvisioningTemplateCreationInformation
    {
        private ProvisioningTemplate baseTemplate;
        private FileConnectorBase fileConnector;
        private bool persistBrandingFiles = false;
        private bool persistMultiLanguageResourceFiles = false;
        private string resourceFilePrefix = "PnP_Resources";
        private bool includeAllTermGroups = false;
        private bool includeSiteCollectionTermGroup = false;
        private bool includeSiteGroups = false;
        private bool includeTermGroupsSecurity = false;
        private bool includeSearchConfiguration = false;
        private List<String> propertyBagPropertiesToPreserve;
        private List<String> contentTypeGroupsToInclude;
        private bool persistPublishingFiles = false;
        private bool includeNativePublishingFiles = false;
        private bool skipVersionCheck = false;
        private List<ExtensibilityHandler> extensibilityHandlers = new List<ExtensibilityHandler>();
        private Handlers handlersToProcess = Handlers.All;
        private bool includeContentTypesFromSyndication = true;
        private bool includeHiddenLists = false;
        private bool includeAllClientSidePages = false;



        /// <summary>
        /// Provisioning Progress Delegate
        /// </summary>
        public ProvisioningProgressDelegate ProgressDelegate { get; set; }

        /// <summary>
        /// Provisioning Messages Delegate
        /// </summary>
        public ProvisioningMessagesDelegate MessagesDelegate { get; set; }

        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="web">A SharePoint site or subsite</param>
        public ProvisioningTemplateCreationInformation(Web web)
        {
            this.baseTemplate = web.GetBaseTemplate();
            this.propertyBagPropertiesToPreserve = new List<String>();
            this.contentTypeGroupsToInclude = new List<String>();
        }

        /// <summary>
        /// Base template used to compare against when we're "getting" a template
        /// </summary>
        public ProvisioningTemplate BaseTemplate
        {
            get
            {
                return this.baseTemplate;
            }
            set
            {
                this.baseTemplate = value;
            }
        }

        /// <summary>
        /// Connector used to persist files when needed
        /// </summary>
        public FileConnectorBase FileConnector
        {
            get
            {
                return this.fileConnector;
            }
            set
            {
                this.fileConnector = value;
            }
        }

        /// <summary>
        /// Will create resource files named "PnP_Resource_[LCID].resx for every supported language. The files will be persisted to the location specified by the connector
        /// </summary>
        public bool PersistMultiLanguageResources
        {
            get
            {
                return this.persistMultiLanguageResourceFiles;
            }
            set
            {
                this.persistMultiLanguageResourceFiles = value;
            }
        }

        /// <summary>
        /// Prefix for resource file
        /// </summary>
        public string ResourceFilePrefix
        {
            get
            {
                return this.resourceFilePrefix;
            }
            set
            {
                this.resourceFilePrefix = value;
            }
        }

        /// <summary>
        /// Do composed look files (theme files, site logo, alternate css) need to be persisted to storage when 
        /// we're "getting" a template
        /// </summary>
        [Obsolete("Use PersistBrandingFiles instead")]
        public bool PersistComposedLookFiles
        {
            get
            {
                return this.persistBrandingFiles;
            }
            set
            {
                this.persistBrandingFiles = value;
            }
        }

        /// <summary>
        /// if true, persists branding files in the template
        /// </summary>
        public bool PersistBrandingFiles
        {
            get
            {
                return this.persistBrandingFiles;
            }
            set
            {
                this.persistBrandingFiles = value;
            }
        }

        /// <summary>
        /// Defines whether to persist publishing files (MasterPages and PageLayouts)
        /// </summary>
        public bool PersistPublishingFiles
        {
            get
            {
                return this.persistPublishingFiles;
            }
            set
            {
                this.persistPublishingFiles = value;
            }
        }

        /// <summary>
        /// Defines whether to extract native publishing files (MasterPages and PageLayouts)
        /// </summary>
        public bool IncludeNativePublishingFiles
        {
            get
            {
                return this.includeNativePublishingFiles;
            }
            set
            {
                this.includeNativePublishingFiles = value;
            }
        }

        /// <summary>
        /// If true includes all term groups in the template
        /// </summary>
        public bool IncludeAllTermGroups
        {
            get
            {
                return this.includeAllTermGroups;
            }
            set { this.includeAllTermGroups = value; }
        }

        /// <summary>
        /// if true, includes site collection term groups in the template
        /// </summary>
        public bool IncludeSiteCollectionTermGroup
        {
            get { return this.includeSiteCollectionTermGroup; }
            set { this.includeSiteCollectionTermGroup = value; }
        }

        /// <summary>
        /// if true, includes term group security in the template
        /// </summary>
        public bool IncludeTermGroupsSecurity
        {
            get { return this.includeTermGroupsSecurity; }
            set { this.includeTermGroupsSecurity = value; }
        }

        internal List<String> PropertyBagPropertiesToPreserve
        {
            get { return this.propertyBagPropertiesToPreserve; }
            set { this.propertyBagPropertiesToPreserve = value; }
        }

        /// <summary>
        /// List of content type groups
        /// </summary>
        public List<String> ContentTypeGroupsToInclude
        {
            get { return this.contentTypeGroupsToInclude; }
            set { this.contentTypeGroupsToInclude = value; }
        }

        /// <summary>
        /// if true, includes site groups in the template
        /// </summary>
        public bool IncludeSiteGroups
        {
            get
            {
                return this.includeSiteGroups;
            }
            set { this.includeSiteGroups = value; }
        }

        /// <summary>
        /// if true includes search configuration in the template
        /// </summary>
        public bool IncludeSearchConfiguration
        {
            get
            {
                return this.includeSearchConfiguration;
            }
            set
            {
                this.includeSearchConfiguration = value;
            }
        }

        /// <summary>
        /// If true all client side pages will be included in the template.
        /// </summary>
        public bool IncludeAllClientSidePages
        {
            get
            {
                return this.includeAllClientSidePages;
            }
            set
            {
                this.includeAllClientSidePages = value;
            }
        }

        /// <summary>
        /// List of of handlers to process
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
        /// if true, skips version check
        /// </summary>
        public bool SkipVersionCheck
        {
            get { return skipVersionCheck; }
            set { skipVersionCheck = value; }
        }

        /// <summary>
        /// Gets or sets a value indicating whether to include content types from syndication (= content type hub) or not.
        /// </summary>
        /// <value>
        ///   <c>true</c> if the export should contains content types issued from syndication (= content type hub)
        /// </value>
        public bool IncludeContentTypesFromSyndication
        {
            get { return includeContentTypesFromSyndication; }
            set { includeContentTypesFromSyndication = value; }
        }

        /// <summary>
        /// Declares whether to include hidden lists in the output or not
        /// </summary>
        public bool IncludeHiddenLists
        {
            get { return includeHiddenLists; }
            set { includeHiddenLists = value; }
        }

        /// <summary>
        /// Optional argument to specify the collection of lists to extract
        /// </summary>
        /// <remarks>
        /// Can contain the title or the ID of the lists to export
        /// </remarks>
        public List<String> ListsToExtract { get; set; } = new List<String>();

        /// <summary>
        /// Specifies whether to also load site collection term groups when initializing the <see cref="TokenParser"/>. If
        /// <c>false</c>, only normal term groups will be loaded. This does not affect loading the site collection term group
        /// when one of the <c>sitecollectionterm</c> tokens was found.
        /// </summary>
        public bool LoadSiteCollectionTermGroups { get; set; } = true;

        /// <summary>
        /// List which contains information about resource tokens used and/or created during the extraction of a template.
        /// </summary>
        internal List<Tuple<string, int, string>> ResourceTokens { get; } = new List<Tuple<string, int, string>>();
        
        /// <summary>
        /// Extraction configuration coming from JSON
        /// </summary>
        internal Model.Configuration.ExtractConfiguration ExtractConfiguration { get; set; }

    }
}
