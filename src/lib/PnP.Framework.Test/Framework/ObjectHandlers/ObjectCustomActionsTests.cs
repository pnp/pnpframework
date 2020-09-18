using Microsoft.SharePoint.Client;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using PnP.Framework.Provisioning.Model;
using PnP.Framework.Provisioning.ObjectHandlers;
using System.Linq;

namespace PnP.Framework.Tests.Framework.ObjectHandlers
{
    [TestClass]
    public class ObjectCustomActionsTests
    {
        private const string ActionName = "Test Custom Action";
        [TestCleanup]
        public void CleanUp()
        {
            using (var ctx = TestCommon.CreateClientContext())
            {
                if (ctx.Site.CustomActionExists("Test Custom Action"))
                {
                    var action = ctx.Site.GetCustomActions().FirstOrDefault(c => c.Name == ActionName);
                    action.DeleteObject();
                    ctx.ExecuteQueryRetry();
                }
            }
        }

        [TestMethod]
        public void CanProvisionObjects()
        {
            var template = new ProvisioningTemplate();
            var ca = new PnP.Framework.Provisioning.Model.CustomAction
            {
                Name = "Test Custom Action",
                Location = "ScriptLink",
                ScriptBlock = "alert('Hello PnP!');"
            };

            template.CustomActions.SiteCustomActions.Add(ca);

            using (var ctx = TestCommon.CreateClientContext())
            {
                var parser = new TokenParser(ctx.Web, template);
                new ObjectCustomActions().ProvisionObjects(ctx.Web, template, parser, new ProvisioningTemplateApplyingInformation());

                Assert.IsTrue(ctx.Site.CustomActionExists("Test Custom Action"));
            }
        }

        [TestMethod]
        public void CanCreateEntities()
        {
            using (var ctx = TestCommon.CreateClientContext())
            {
                // Load the base template which will be used for the comparison work
                var creationInfo = new ProvisioningTemplateCreationInformation(ctx.Web) { BaseTemplate = ctx.Web.GetBaseTemplate() };

                var template = new ProvisioningTemplate();
                template = new ObjectCustomActions().ExtractObjects(ctx.Web, template, creationInfo);

                Assert.IsInstanceOfType(template.CustomActions, typeof(CustomActions));
            }
        }
    }
}
