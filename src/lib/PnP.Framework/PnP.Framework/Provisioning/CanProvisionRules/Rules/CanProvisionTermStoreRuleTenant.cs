using Microsoft.Online.SharePoint.TenantAdministration;
using PnP.Framework.Provisioning.Model;
using PnP.Framework.Provisioning.ObjectHandlers;

namespace PnP.Framework.Provisioning.CanProvisionRules.Rules
{
    [CanProvisionRule(Scope = CanProvisionScope.Tenant, Sequence = 200)]
    internal class CanProvisionTermStoreRuleTenant : CanProvisionRuleTenantBase
    {
        public override CanProvisionResult CanProvision(Tenant tenant, ProvisioningHierarchy hierarchy, string sequenceId, ProvisioningTemplateApplyingInformation applyingInformation)
        {
            // Rely on the corresponding Site level CanProvision rule
            return (this.EvaluateSiteRule<CanProvisionTermStoreRuleSite>(tenant, hierarchy, sequenceId, applyingInformation));
        }
    }
}
