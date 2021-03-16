using Microsoft.VisualStudio.TestTools.UnitTesting;
using PnP.Framework.ALM;
using PnP.Framework.Test.Properties;
using System;
using System.Linq;
using System.Threading.Tasks;

namespace PnP.Framework.Test.Sites
{
    [TestClass]
    public class ALMTests
    {
        private Guid appGuid;

        [TestInitialize]
        public void Initialize()
        {
            appGuid = Guid.Parse("f11fc8ba-e64b-4354-b227-fded0bb31cfc");

        }

        [TestCleanup]
        public void CleanUp()
        {

        }

        [TestMethod]
        public async Task AddCheckRemoveAppTestAsync()
        {
            TestCommon.RegisterPnPHttpClientMock();
            using (var clientContext = TestCommon.CreateTestClientContext())
            {
                AppManager manager = new AppManager(clientContext);
                var appBytes = Resources.alm;

                //Test adding app
                var addedApp = await manager.AddAsync(appBytes, $"app-{appGuid}.sppkg", true);

                Assert.IsNotNull(addedApp);


                //Test availability of apps
                var availableApps = await manager.GetAvailableAsync();

                Assert.IsNotNull(availableApps);
                CollectionAssert.Contains(availableApps.Select(app => app.Id).ToList(), addedApp.Id);

                var retrievedApp = await manager.GetAvailableAsync(addedApp.Id);
                Assert.AreEqual(addedApp.Id, retrievedApp.Id);

                //Test removal
                var removeResults = await manager.RemoveAsync(addedApp.Id);

                Assert.IsTrue(removeResults);
            }
        }

        [TestMethod]
        public void AddCheckRemoveAppTest()
        {
            TestCommon.RegisterPnPHttpClientMock();
            using (var clientContext = TestCommon.CreateTestClientContext())
            {
                AppManager manager = new AppManager(clientContext);
                var appBytes = Resources.alm;

                //Test adding app
                var addedApp = manager.Add(appBytes, $"app-{appGuid}.sppkg", true);

                Assert.IsNotNull(addedApp);

                //Test availability of apps
                var availableApps = manager.GetAvailable();

                Assert.IsNotNull(availableApps);
                CollectionAssert.Contains(availableApps.Select(app => app.Id).ToList(), addedApp.Id);

                var retrievedApp = manager.GetAvailable(addedApp.Id);
                Assert.AreEqual(addedApp.Id, retrievedApp.Id);

                //Test removal
                var removeResults = manager.Remove(addedApp.Id);

                Assert.IsTrue(removeResults);
            }
        }

        [TestMethod]
        public void DeployRetractAppTest()
        {
            TestCommon.RegisterPnPHttpClientMock();
            using (var clientContext = TestCommon.CreateTestClientContext())
            {
                AppManager manager = new AppManager(clientContext);
                var appBytes = Resources.almskip;

                var results = manager.Add(appBytes, $"appalmskip-{appGuid}.sppkg", true);

                var deployResults = manager.Deploy(results.Id);
                Assert.IsTrue(deployResults);

                var metadata = manager.GetAvailable(results.Id);
                Assert.IsTrue(metadata.Deployed);

                var retractResults = manager.Retract(results.Id);
                Assert.IsTrue(retractResults);

                metadata = manager.GetAvailable(results.Id);
                Assert.IsFalse(metadata.Deployed);

                manager.Remove(results.Id);
            }
        }

        [TestMethod]
        public async Task DeployRetractAppAsyncTest()
        {
            TestCommon.RegisterPnPHttpClientMock();
            using (var clientContext = TestCommon.CreateTestClientContext())
            {
                AppManager manager = new AppManager(clientContext);
                var appBytes = Resources.almskip;

                var results = await manager.AddAsync(appBytes, $"appalmskip-{appGuid}.sppkg", true);

                var deployResults = await manager.DeployAsync(results.Id);
                Assert.IsTrue(deployResults);

                var metadata = await manager.GetAvailableAsync(results.Id);
                Assert.IsTrue(metadata.Deployed);

                var retractResults = await manager.RetractAsync(results.Id);
                Assert.IsTrue(retractResults);

                metadata = await manager.GetAvailableAsync(results.Id);
                Assert.IsFalse(metadata.Deployed);

                manager.Remove(results.Id);
            }
        }

    }
}
