using PnP.Framework.Provisioning.ObjectHandlers;

namespace PnP.Framework.Provisioning.CanProvisionRules
{
    /// <summary>
    /// Interface to test if a template can be provisioned onto a target whole Office 365 tenant
    /// </summary>
    internal interface ICanProvisionRuleOffice365 : ICanProvisionRuleBase
    {
        /// <summary>
        /// This method allows to check if a template can be provisioned
        /// </summary>
        /// <param name="hierarchy">The Template to hierarchy</param>
        /// <param name="sequenceId">The sequence to test within the hierarchy</param>
        /// <param name="applyingInformation">Any custom provisioning settings</param>
        /// <returns>A boolean stating whether the current object handler can be run during provisioning or if there are any missing requirements</returns>
        CanProvisionResult CanProvision(Model.ProvisioningHierarchy hierarchy, string sequenceId, ProvisioningTemplateApplyingInformation applyingInformation);
    }
}
