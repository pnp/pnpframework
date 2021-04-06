using Microsoft.SharePoint.Client;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using PnP.Framework.Provisioning.Model;
using PnP.Framework.Provisioning.ObjectHandlers;
using PnP.Framework.Provisioning.Providers.Xml;
using PnP.Framework.Utilities;
using System;

namespace PnP.Framework.Test.Framework.ProvisioningTemplates
{
    [TestClass]
    public class ClientSidePageProvisioningTests
    {

        [TestCleanup]
        public void Cleanup()
        {
            using (var ctx = TestCommon.CreateTestClientContext())
            {
                ctx.Load(ctx.Web, w => w.ServerRelativeUrl);
                ctx.ExecuteQueryRetry();

                TestCommon.DeleteFile(ctx, UrlUtility.Combine(ctx.Web.ServerRelativeUrl, "/SitePages/csp-test-1.aspx"));
                TestCommon.DeleteFile(ctx, UrlUtility.Combine(ctx.Web.ServerRelativeUrl, "/SitePages/csp-test-2.aspx"));
                TestCommon.DeleteFile(ctx, UrlUtility.Combine(ctx.Web.ServerRelativeUrl, "/SitePages/csp-test-3.aspx"));
            }
        }

        // background for this test: https://github.com/SharePoint/PnP-Sites-Core/issues/2203
        [TestMethod]
        public void ProvisionClientSidePagesWithHeader()
        {
            var resourceFolder = string.Format(@"{0}\Resources\Templates", AppDomain.CurrentDomain.BaseDirectory);
            XMLTemplateProvider provider = new XMLFileSystemTemplateProvider(resourceFolder, "");

            var existingTemplate = provider.GetTemplate("ClientSidePagesWithHeader.xml");
            using (var ctx = TestCommon.CreateTestClientContext())
            {
                ctx.Web.ApplyProvisioningTemplate(existingTemplate, new ProvisioningTemplateApplyingInformation()
                {
                    HandlersToProcess = Handlers.Pages
                });
            }
        }
    }
}
