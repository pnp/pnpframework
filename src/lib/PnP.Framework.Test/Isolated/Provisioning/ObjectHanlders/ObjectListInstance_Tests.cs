using Microsoft.SharePoint.Client;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using PnP.Framework.Provisioning.Model;
using PnP.Framework.Provisioning.ObjectHandlers;
using PnP.Framework.Utilities;
using PnP.Framework.Utilities.UnitTests.Helpers;
using PnP.Framework.Utilities.UnitTests.Model;
using PnP.Framework.Utilities.UnitTests.Web;
using System;
using System.Collections.Generic;
using System.Text;

namespace PnP.Framework.Test.Isolated.Provisioning.ObjectHanlders
{
    [TestClass]
    public class ObjectListInstance_Tests
    {
        [TestMethod]
        public void ObjectListInstance_ProvisionObjects()
        {
            ObjectListInstance handler = new ObjectListInstance(FieldAndListProvisioningStepHelper.Step.ListAndStandardFields);
            ProvisioningTemplate template = new ProvisioningTemplate();
            string listName = "PnP Unit Test List";
            var listInstance = new ListInstance
            {
                Url = string.Format("lists/{0}", listName),
                Title = listName,
                TemplateType = (int)ListTemplateType.GenericList
            };
            listInstance.FieldRefs.Add(new FieldRef() { Id = new Guid("23f27201-bee3-471e-b2e7-b64fd8b7ca38") });
            template.Lists.Add(listInstance);

            using (var ctx = TestCommon.CreateTestClientContext(false))
            {
                TokenParser parser = new TokenParser(ctx.Web, template);
                handler.ProvisionObjects(ctx.Web, template, parser, null);
            }
        }
    }
}
