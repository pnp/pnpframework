using Microsoft.SharePoint.Client;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using PnP.Framework.Modernization.Publishing;

namespace PnP.Framework.Modernization.Tests.Transform.Publishing
{
    [TestClass]
    public class AdHocTests
    {

        [TestMethod]
        public void TestMethod1()
        {
            using (ClientContext cc = TestCommon.CreateClientContext())
            {
                PageLayoutManager m = new PageLayoutManager(null);
                var result = m.LoadPageLayoutMappingFile(@"..\..\..\PnP.Framework.Modernization\Publishing\pagelayoutmapping_sample.xml");
            }
        }

    }
}
