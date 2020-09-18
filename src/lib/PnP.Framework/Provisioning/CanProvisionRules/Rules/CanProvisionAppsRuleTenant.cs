using Microsoft.Online.SharePoint.TenantAdministration;
using PnP.Framework.Provisioning.Model;
using PnP.Framework.Provisioning.ObjectHandlers;

namespace PnP.Framework.Provisioning.CanProvisionRules.Rules
{
    [CanProvisionRule(Scope = CanProvisionScope.Tenant, Sequence = 100)]
    internal class CanProvisionAppsRuleTenant : CanProvisionRuleTenantBase
    {
        public override CanProvisionResult CanProvision(Tenant tenant, ProvisioningHierarchy hierarchy, string sequenceId, ProvisioningTemplateApplyingInformation applyingInformation)
        {
            // Rely on the corresponding Site level CanProvision rule
            return (this.EvaluateSiteRule<CanProvisionAppsRuleSite>(tenant, hierarchy, sequenceId, applyingInformation));
        }
    }
}
