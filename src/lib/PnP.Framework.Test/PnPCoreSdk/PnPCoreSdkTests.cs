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

    }
}
