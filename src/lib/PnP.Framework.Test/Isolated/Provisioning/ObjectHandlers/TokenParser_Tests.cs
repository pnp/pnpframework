using Microsoft.SharePoint.Client;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using PnP.Framework.Provisioning.Model;
using PnP.Framework.Provisioning.ObjectHandlers;
using PnP.Framework.Utilities.UnitTests.Model;
using PnP.Framework.Utilities.UnitTests.Web;
using System;
using System.Collections.Generic;
using System.Text;

namespace PnP.Framework.Test.Isolated.Provisioning.ObjectHanlders
{
    [TestClass]
    public class TokenParser_Tests
    {
        [TestMethod]
        public void TokenParser_Test_ParseStringWebPart()
        {

            string mockSiteUrl = "https://test.sharepoint.com/sites/TestWeb";
            ProvisioningTemplate template = new ProvisioningTemplate();
            template.Lists.Add(new ListInstance()
            {
                Title = "{sitetitle}"
            });
            using (ClientContext cctx = new ClientContext(mockSiteUrl))
            {
                MockEntryResponseProvider responseProvider = new MockEntryResponseProvider();
                responseProvider.ResponseEntries.Add(new MockResponseEntry<object>()
                {
                    Url = mockSiteUrl,
                    PropertyName = "Web",
                    ReturnValue = new
                    {
                        Title = "Test Web",
                        ServerRelativeUrl = "/sites/TestWeb",
                        Url = mockSiteUrl
                    }
                });
                MockWebRequestExecutorFactory executorFactory = new MockWebRequestExecutorFactory(responseProvider);
                cctx.WebRequestExecutorFactory = executorFactory;

                TokenParser parser = new TokenParser(cctx.Web, template);
                Assert.AreEqual("Test Web", parser.ParseString("{sitetitle}"));
            }
        }
    }
}
