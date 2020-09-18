using Microsoft.SharePoint.Client;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using PnP.Framework.Provisioning.Model;
using PnP.Framework.Provisioning.ObjectHandlers;
using System;
using System.Linq;

namespace PnP.Framework.Test.Framework.ObjectHandlers
{
    [TestClass]
    public class ObjectPropertyBagEntryTests
    {
        private string key;
        private string systemKey;

        [TestInitialize]
        public void Initialize()
        {
            key = string.Format("Test_{0}", DateTime.Now.Ticks);
            systemKey = string.Format("vti_test_{0}", DateTime.Now.Ticks);
        }
        [TestCleanup]
        public void CleanUp()
        {
            using (var ctx = TestCommon.CreateClientContext())
            {
                ctx.Web.RemovePropertyBagValue(key);
                ctx.Web.RemovePropertyBagValue(systemKey);
            }
        }

        [TestMethod]
        public void CanProvisionObjects()
        {
            var template = new ProvisioningTemplate();

            var propbagEntry = new PnP.Framework.Provisioning.Model.PropertyBagEntry
            {
                Key = key,
                Value = "Unit Test"
            };

            template.PropertyBagEntries.Add(propbagEntry);

            using (var ctx = TestCommon.CreateClientContext())
            {
                var parser = new TokenParser(ctx.Web, template);
                parser = new ObjectPropertyBagEntry().ProvisionObjects(ctx.Web, template, parser, new ProvisioningTemplateApplyingInformation());

                var value = ctx.Web.GetPropertyBagValueString(key, "default");
                Assert.IsTrue(value == "Unit Test");

                // Create same entry, but don't overwrite.
                template = new ProvisioningTemplate();

                var propbagEntry2 = new PropertyBagEntry
                {
                    Key = key,
                    Value = "Unit Test 2",
                    Overwrite = false
                };

                template.PropertyBagEntries.Add(propbagEntry2);

                parser = new ObjectPropertyBagEntry().ProvisionObjects(ctx.Web, template, parser, new ProvisioningTemplateApplyingInformation());

                value = ctx.Web.GetPropertyBagValueString(key, "default");
                Assert.IsTrue(value == "Unit Test");


                // Create same entry, but overwrite
                template = new ProvisioningTemplate();

                var propbagEntry3 = new PropertyBagEntry
                {
                    Key = key,
                    Value = "Unit Test 3",
                    Overwrite = true
                };

                template.PropertyBagEntries.Add(propbagEntry3);

                parser = new ObjectPropertyBagEntry().ProvisionObjects(ctx.Web, template, parser, new ProvisioningTemplateApplyingInformation());

                value = ctx.Web.GetPropertyBagValueString(key, "default");
                Assert.IsTrue(value == "Unit Test 3");

                // Create entry with system key. We don't specify to overwrite system keys, so the key should not be created.
                template = new ProvisioningTemplate();

                var propbagEntry4 = new PropertyBagEntry
                {
                    Key = systemKey,
                    Value = "Unit Test System Key",
                    Overwrite = true
                };

                template.PropertyBagEntries.Add(propbagEntry4);

                parser = new ObjectPropertyBagEntry().ProvisionObjects(ctx.Web, template, parser, new ProvisioningTemplateApplyingInformation());

                value = ctx.Web.GetPropertyBagValueString(systemKey, "default");
                Assert.IsTrue(value == "default");

                // Create entry with system key. We _do_ specify to overwrite system keys, so the key should be created.
                template = new ProvisioningTemplate();

                var propbagEntry5 = new PropertyBagEntry
                {
                    Key = systemKey,
                    Value = "Unit Test System Key 5",
                    Overwrite = true
                };

                template.PropertyBagEntries.Add(propbagEntry5);

                parser = new ObjectPropertyBagEntry().ProvisionObjects(ctx.Web, template, parser, new ProvisioningTemplateApplyingInformation() { OverwriteSystemPropertyBagValues = true });

                value = ctx.Web.GetPropertyBagValueString(systemKey, "default");
                Assert.IsTrue(value == "Unit Test System Key 5");

                // Create entry with system key. We _do not_ specify to overwrite system keys, so the key should not be created.
                template = new ProvisioningTemplate();

                var propbagEntry6 = new PropertyBagEntry
                {
                    Key = systemKey,
                    Value = "Unit Test System Key 6",
                    Overwrite = true
                };

                template.PropertyBagEntries.Add(propbagEntry6);

                parser = new ObjectPropertyBagEntry().ProvisionObjects(ctx.Web, template, parser, new ProvisioningTemplateApplyingInformation() { OverwriteSystemPropertyBagValues = false });

                value = ctx.Web.GetPropertyBagValueString(systemKey, "default");
                Assert.IsFalse(value == "Unit Test System Key 6");
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
                template = new ObjectPropertyBagEntry().ExtractObjects(ctx.Web, template, creationInfo);

                Assert.IsTrue(template.PropertyBagEntries.Any());
            }
        }
    }
}
