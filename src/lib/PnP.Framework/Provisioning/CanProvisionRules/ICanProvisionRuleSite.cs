using Microsoft.SharePoint.Client;
using PnP.Framework.Provisioning.ObjectHandlers;

namespace PnP.Framework.Provisioning.CanProvisionRules
{
    /// <summary>
    /// Interface to test if a template can be provisioned onto a target site
    /// </summary>
    internal interface ICanProvisionRuleSite : ICanProvisionRuleBase
    {
        /// <summary>
        /// This method allows to check if a template can be provisioned in the currently selected target
        /// </summary>
        /// <param name="web">The target Web</param>
        /// <param name="template">The Template to provision</param>
        /// <param name="applyingInformation">Any custom provisioning settings</param>
        /// <returns>A boolean stating whether the current object handler can be run during provisioning or if there are any missing requirements</returns>
        CanProvisionResult CanProvision(Web web, Model.ProvisioningTemplate template, ProvisioningTemplateApplyingInformation applyingInformation);
    }
}
