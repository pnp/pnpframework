using Microsoft.SharePoint.Client;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using PnP.Framework.Provisioning.Model;
using PnP.Framework.Provisioning.ObjectHandlers;

namespace PnP.Framework.Tests.Framework.ObjectHandlers
{
    [TestClass]
    public class ObjectComposedLookTests
    {

        [TestMethod]
        public void CanCreateComposedLooks()
        {
            using (var scope = new PnP.Framework.Diagnostics.PnPMonitoredScope("ComposedLookTests"))
            {
                using (var ctx = TestCommon.CreateClientContext())
                {
                    // Load the base template which will be used for the comparison work
                    var creationInfo = new ProvisioningTemplateCreationInformation(ctx.Web) { BaseTemplate = ctx.Web.GetBaseTemplate() };

                    var template = new ProvisioningTemplate();
                    template = new ObjectComposedLook().ExtractObjects(ctx.Web, template, creationInfo);
                    Assert.IsInstanceOfType(template.ComposedLook, typeof(PnP.Framework.Provisioning.Model.ComposedLook));
                }
            }
        }
    }
}
