using Microsoft.Online.SharePoint.TenantAdministration;
using Microsoft.SharePoint.Client;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using PnP.Framework.ALM;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace PnP.Framework.Tests.Sites
{
    [TestClass]
    public class ALMTests
    {
        private Guid appGuid;

        [TestInitialize]
        public void Initialize()
        {
            appGuid = Guid.NewGuid();

        }

        [TestCleanup]
        public void CleanUp()
        {

        }

        [TestMethod]
        public async Task AddCheckRemoveAppTestAsync()
        {
            using (var clientContext = TestCommon.CreateClientContext())
            {
                AppManager manager = new AppManager(clientContext);
                var appBytes = PnP.Framework.Tests.Properties.Resources.alm;

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
            using (var clientContext = TestCommon.CreateClientContext())
            {
                AppManager manager = new AppManager(clientContext);
                var appBytes = PnP.Framework.Tests.Properties.Resources.alm;

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
            using (var clientContext = TestCommon.CreateClientContext())
            {
                AppManager manager = new AppManager(clientContext);
                var appBytes = PnP.Framework.Tests.Properties.Resources.almskip;

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
            using (var clientContext = TestCommon.CreateClientContext())
            {
                AppManager manager = new AppManager(clientContext);
                var appBytes = PnP.Framework.Tests.Properties.Resources.almskip;

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
