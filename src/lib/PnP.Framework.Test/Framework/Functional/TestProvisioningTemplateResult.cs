using PnP.Framework.Provisioning.Model;
using PnP.Framework.Provisioning.ObjectHandlers;

namespace PnP.Framework.Tests.Framework.Functional
{
    public class TestProvisioningTemplateResult
    {
        public ProvisioningTemplate SourceTemplate { get; set; }
        public TokenParser SourceTokenParser { get; set; }
        public ProvisioningTemplate TargetTemplate { get; set; }
        public TokenParser TargetTokenParser { get; set; }

    }
}
