using Microsoft.Online.SharePoint.TenantAdministration;
using Microsoft.SharePoint.Client;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using PnP.Framework.Entities;
using PnP.Framework.Provisioning.Model;
using PnP.Framework.Provisioning.ObjectHandlers;
using PnP.Framework.Provisioning.ObjectHandlers.TokenDefinitions;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using User = PnP.Framework.Provisioning.Model.User;

namespace PnP.Framework.Test.Framework.ObjectHandlers
{
    [TestClass]
    public class ObjectSiteSettingsTests
    {
        [TestInitialize]
        public void Initialize()
        {
            TestCommon.FixAssemblyResolving("Newtonsoft.Json");
        }

        [TestCleanup]
        public void CleanUp()
        {
           
        }

        [TestMethod]
        public void CanProvisionObjects_ShowPeoplePickerSuggestionsForGuestUsers()
        {
            // Ensure Clean Start
            using (var ctxClean = TestCommon.CreateClientContext())
            {
                ctxClean.Load(ctxClean.Site, p => p.ShowPeoplePickerSuggestionsForGuestUsers);
                ctxClean.ExecuteQueryRetry();
                
                ctxClean.Site.ShowPeoplePickerSuggestionsForGuestUsers = false;
                ctxClean.ExecuteQueryRetry();
            }

            // Set On
            var templateOn = new ProvisioningTemplate();
            templateOn.SiteSettings = new SiteSettings() { ShowPeoplePickerSuggestionsForGuestUsers = true };

            using (var ctx = TestCommon.CreateClientContext())
            {
                var parser = new TokenParser(ctx.Web, templateOn);
                new ObjectSiteSettings().ProvisionObjects(ctx.Web, templateOn, parser, new ProvisioningTemplateApplyingInformation());
                ctx.ExecuteQueryRetry();
            }

            //Check On
            using (var ctxCheckOn = TestCommon.CreateClientContext())
            {
                ctxCheckOn.Load(ctxCheckOn.Site, p => p.ShowPeoplePickerSuggestionsForGuestUsers);
                ctxCheckOn.ExecuteQueryRetry();

                Assert.IsTrue(ctxCheckOn.Site.ShowPeoplePickerSuggestionsForGuestUsers);
            }

            //Set Off
            var templateOff = new ProvisioningTemplate();
            templateOff.SiteSettings = new SiteSettings() { ShowPeoplePickerSuggestionsForGuestUsers = false };

            using (var ctxOff = TestCommon.CreateClientContext())
            {
                var parser = new TokenParser(ctxOff.Web, templateOff);
                new ObjectSiteSettings().ProvisionObjects(ctxOff.Web, templateOff, parser, new ProvisioningTemplateApplyingInformation());
                ctxOff.ExecuteQueryRetry();
            }
                       
            //Check Off
            using (var ctxCheckOff = TestCommon.CreateClientContext())
            {
                ctxCheckOff.Load(ctxCheckOff.Site, p => p.ShowPeoplePickerSuggestionsForGuestUsers);
                ctxCheckOff.ExecuteQueryRetry();

                Assert.IsFalse(ctxCheckOff.Site.ShowPeoplePickerSuggestionsForGuestUsers);
            }

        }
       
    }
}
