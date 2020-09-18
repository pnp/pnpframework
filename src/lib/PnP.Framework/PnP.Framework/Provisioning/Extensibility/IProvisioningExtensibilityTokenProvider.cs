using Microsoft.SharePoint.Client;
using PnP.Framework.Provisioning.Model;
using PnP.Framework.Provisioning.ObjectHandlers.TokenDefinitions;
using System.Collections.Generic;

namespace PnP.Framework.Provisioning.Extensibility
{
    /// <summary>
    /// Defines an interface which allows to plugin custom TokenDefinitions to the template provisioning pipeline
    /// </summary>
    public interface IProvisioningExtensibilityTokenProvider
    {
        /// <summary>
        /// Provides Token Definitions to the template provisioning pipeline
        /// </summary>
        /// <param name="ctx">The ClientContext</param>
        /// <param name="template">The Provisioning template</param>
        /// <param name="configurationData">Configuration Data string</param>
        IEnumerable<TokenDefinition> GetTokens(ClientContext ctx, ProvisioningTemplate template, string configurationData);
    }
}
