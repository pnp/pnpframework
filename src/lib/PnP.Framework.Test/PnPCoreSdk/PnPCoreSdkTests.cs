using Microsoft.SharePoint.Client;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using PnP.Core.Model;
using PnP.Core.Services;

namespace PnP.Framework.Test
{
    [TestClass]
    public class PnPCoreSdkTests
    {
        #region Test initialization
        [ClassInitialize()]
        public static void ClassInit(TestContext context)
        {
        }

        [ClassCleanup()]
        public static void ClassCleanup()
        {
        }
        #endregion

        [TestMethod]
        public void GetWebTest()
        {
            using (var cc = TestCommon.CreateClientContext())
            {
                using (PnPContext context = PnPCoreSdk.Instance.GetPnPContext(cc))
                {
                    var web = context.Web.GetAsync().GetAwaiter().GetResult();                                    
                    Assert.IsTrue(web.Title != null);
                }
            }
        }

        [TestMethod]
        public void GetWebReverseTest()
        {
            using (var cc = TestCommon.CreateClientContext())
            {
                using (PnPContext context = PnPCoreSdk.Instance.GetPnPContext(cc))
                {                    
                    using (ClientContext ccAgain = PnPCoreSdk.Instance.GetClientContext(context))
                    {
                        ccAgain.Load(ccAgain.Web, p => p.Title);
                        ccAgain.ExecuteQueryRetry();
                        Assert.IsTrue(ccAgain.Web.Title != null);
                    }
                }
            }
        }

        [TestMethod]
        public void PassAzureEnvironmentTest()
        {
            using (var cc = TestCommon.CreateClientContext(AzureEnvironment.Custom))
            {
                using (PnPContext context = PnPCoreSdk.Instance.GetPnPContext(cc))
                {
                    var web = context.Web.GetAsync().GetAwaiter().GetResult();
                    Assert.IsTrue(web.Title != null);
                }
            }
        }

        [TestMethod]
        public void PassAzureEnvironmentReverseTest()
        {
            using (var cc = TestCommon.CreateClientContext(AzureEnvironment.Custom))
            {
                using (PnPContext context = PnPCoreSdk.Instance.GetPnPContext(cc))
                {
                    using (ClientContext ccAgain = PnPCoreSdk.Instance.GetClientContext(context))
                    {
                        ccAgain.Load(ccAgain.Web, p => p.Title);
                        ccAgain.ExecuteQueryRetry();
                        Assert.IsTrue(ccAgain.Web.Title != null);
                    }
                }
            }
        }

    }
}
