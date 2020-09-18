using Microsoft.VisualStudio.TestTools.UnitTesting;
using PnP.Framework.Provisioning.Model;
using PnP.Framework.Provisioning.ObjectHandlers;
using PnP.Framework.Test.Framework.Functional.Validators;

namespace PnP.Framework.Test.Framework.Functional.Implementation
{
    internal class SecurityImplementation : ImplementationBase
    {
        internal void SiteCollectionSecurity(string url)
        {
            using (var cc = TestCommon.CreateClientContext(url))
            {
                ProvisioningTemplateCreationInformation ptci = new ProvisioningTemplateCreationInformation(cc.Web)
                {
                    IncludeSiteGroups = true,
                    HandlersToProcess = Handlers.SiteSecurity
                };

                var result = TestProvisioningTemplate(cc, "security_add.xml", Handlers.SiteSecurity, null, ptci);
                SecurityValidator sv = new SecurityValidator();
                Assert.IsTrue(sv.Validate(result.SourceTemplate.Security, result.TargetTemplate.Security, result.TargetTokenParser, cc));
            }
        }
    }
}