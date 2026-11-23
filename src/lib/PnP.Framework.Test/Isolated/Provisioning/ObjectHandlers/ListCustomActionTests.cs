using Microsoft.SharePoint.Client;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using PnP.Framework.Diagnostics;
using PnP.Framework.Provisioning.Model;
using PnP.Framework.Provisioning.ObjectHandlers;
using PnP.Framework.Utilities.UnitTests.Model;
using PnP.Framework.Utilities.UnitTests.Web;
using System;
using System.Reflection;

namespace PnP.Framework.Test.Isolated.Provisioning.ObjectHandlers
{
    [TestClass]
    public class ListCustomActionTests
    {
        [DataTestMethod]
        [DataRow(false, false, true, null, 0, false)]
        [DataRow(true, false, true, null, 0, false)]
        [DataRow(true, true, true, null, 1, true)]
        [DataRow(true, true, false, null, 1, false)]
        [DataRow(true, true, false, false, 1, false)]
        [DataRow(true, true, false, true, 1, true)]
        public void NoScriptListActionsHonorRemovalAndEnabledWhenRightsAreOmitted(
            bool hasAction, bool isSpfx, bool targetExists, bool? enabledForCreation, int expectedRequests, bool expectedChange)
        {
            const string siteUrl = "https://custom-action-tests.sharepoint.invalid/sites/test";
            var responses = new CountingResponseProvider();
            responses.Reads.ResponseEntries.Add(new MockResponseEntry
            {
                Url = siteUrl,
                PropertyName = "UserCustomActions",
                ReturnValue = new
                {
                    _Child_Items_ = targetExists
                        ? new[] { new { _ObjectType_ = "SP.UserCustomAction", _ObjectIdentity_ = "action", Name = "TestAction" } }
                        : Array.Empty<object>()
                }
            });
            using var context = new ClientContext(siteUrl)
            {
                WebRequestExecutorFactory = new MockWebRequestExecutorFactory(responses)
            };
            var templateList = new ListInstance();
            if (hasAction)
            {
                templateList.UserCustomActions.Add(new CustomAction
                {
                    Name = "TestAction",
                    Location = "ClientSideExtension.ListViewCommandSet.CommandBar",
                    Remove = !enabledForCreation.HasValue,
                    Enabled = enabledForCreation ?? true,
                    ClientSideComponentId = isSpfx ? Guid.NewGuid() : Guid.Empty
                });
            }

            using var scope = new PnPMonitoredScope("Custom action test");
            var changed = (bool)typeof(ObjectListInstance)
                .GetMethod("UpdateCustomActions", BindingFlags.Static | BindingFlags.NonPublic)
                .Invoke(null, new object[] { context.Web, context.Web.Lists.GetByTitle("Documents"), templateList, null, scope, true });

            Assert.AreEqual(expectedRequests, responses.Requests);
            Assert.AreEqual(expectedChange, changed);
            if (hasAction && isSpfx && !targetExists)
            {
                Assert.AreEqual(expectedChange, context.HasPendingRequest, "Only an enabled creation should queue an Add.");
            }
        }

        private sealed class CountingResponseProvider : IMockResponseProvider
        {
            public MockEntryResponseProvider Reads { get; } = new MockEntryResponseProvider();
            public int Requests { get; private set; }

            public string GetResponse(string url, string verb, string body)
            {
                Requests++;
                return Reads.GetResponse(url, verb, body);
            }
        }
    }
}
