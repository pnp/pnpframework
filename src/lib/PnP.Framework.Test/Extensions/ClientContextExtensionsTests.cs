using Microsoft.Online.SharePoint.TenantAdministration;
using Microsoft.SharePoint.Client;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Net;

namespace PnP.Framework.Test.Extensions
{
    [TestClass]
    public class ClientContextExtensionsTests
    {
        [DataTestMethod]
        [DataRow("https://tenant.example:443/sites/other", true)]
        [DataRow("https://other.example/sites/team", false)]
        [DataRow("https://tenant.example:444/sites/team", false)]
        [DataRow("http://tenant.example/sites/team", false)]
        public void CloneCopiesExplicitCredentialsOnlyWithinTheSameOrigin(string targetUrl, bool sameOrigin)
        {
            using var source = new ClientContext("https://tenant.example/sites/team");
            var credentials = new NetworkCredential("test-user", "test-password");
            source.Credentials = credentials;

            using var clone = source.Clone(targetUrl);

            Assert.AreSame(sameOrigin ? credentials : null, clone.Credentials);
        }

        [TestMethod]
        public void GetAzureEnvironmentFromContext()
        {
            using (ClientContext clientContext = TestCommon.CreateClientContext())
            {
                Assert.IsTrue(clientContext.GetAzureEnvironment() == AzureEnvironment.Production);
            }
        }

        [TestMethod]
        public void GetAzureEnvironmentFromManualContext()
        {
            using (ClientContext clientContext = new ClientContext("https://contoso.sharepoint.us/sites/bla"))
            {
                Assert.IsTrue(clientContext.GetAzureEnvironment() == AzureEnvironment.USGovernment);
            }
        }

        [TestMethod]
        public void GetAzureEnvironmentFromClonedContext()
        {
            using (ClientContext clientContext = TestCommon.CreateClientContext())
            {
                using (var clientContextCloned = clientContext.Clone("https://contoso.sharepoint.com"))
                {
                    Assert.IsTrue(clientContextCloned.GetAzureEnvironment() == AzureEnvironment.Production);
                }
            }
        }

        [TestMethod]
        public void GetAzureEnvironmentFromTenantContext()
        {
            using (var cc = TestCommon.CreateTenantClientContext())
            {
                Tenant tenant = new Tenant(cc);
                var tenantInstances = tenant.GetTenantInstances();
                cc.Load(tenantInstances);
                cc.ExecuteQuery();

                Assert.IsTrue(tenant.Context.GetAzureEnvironment() == AzureEnvironment.Production);
            }
        }
    }
}
