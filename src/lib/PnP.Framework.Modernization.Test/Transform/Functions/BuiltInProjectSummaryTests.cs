using Microsoft.SharePoint.Client;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Moq;
using Newtonsoft.Json.Linq;
using PnP.Framework.Modernization.Cache;
using PnP.Framework.Modernization.Entities;
using PnP.Framework.Modernization.Functions;
using PnP.Framework.Modernization.Pages;
using PnP.Framework.Modernization.Telemetry;
using PnP.Framework.Modernization.Transform;
using PnP.Framework.Utilities.UnitTests.Model;
using PnP.Framework.Utilities.UnitTests.Web;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using PnPCore = PnP.Core.Model.SharePoint;

namespace PnP.Framework.Modernization.Test.Transform.Functions
{
    [TestClass]
    public class BuiltInProjectSummaryTests
    {
        private const string SourceUrl = "https://contoso.sharepoint.com/sites/source";
        private const string TargetUrl = "https://contoso.sharepoint.com/sites/target";
        private static readonly Guid SourceId = new Guid("73920e94-314e-41ef-820c-d37588137c2b");
        private static readonly Guid TargetId = new Guid("7187a8cb-616c-49ee-a49f-184f8d2e4443");
        private const string Panels = "[{\"Enabled\":true,\"Title\":\"Upcoming \\\"seven\\\"\",\"IsLatePanel\":false,\"DaysToSearch\":7,\"ListUrls\":[\"Lists/Tasks\"]}]";

        [DataTestMethod]
        [DataRow(false, false)]
        [DataRow(false, true)]
        [DataRow(true, false)]
        [DataRow(true, true)]
        public void Mapping_UsesTargetIdentityExportedPanelsAndTargetTimeline(bool crossSite, bool showTimeline)
        {
            using var source = CreateContext(SourceUrl, SourceId, 1);
            using var target = CreateContext(crossSite ? TargetUrl : SourceUrl, crossSite ? TargetId : SourceId, crossSite ? 9 : 1);
            var transformation = BasePageTransformator.LoadDefaultWebPartMapping();
            var mapping = transformation.WebParts.Single(part => part.Type == WebParts.ProjectSummary);
            Assert.IsTrue(mapping.CrossSiteTransformationSupported);
            mapping.Properties = mapping.Properties.Concat(new[]
            {
                transformation.BaseWebPart.Properties.Single(property => property.Name == "Title")
            }).ToArray();
            var analyser = new WikiPage(source.Web.Lists.GetByTitle("SitePages").GetItemById(1), transformation);
            var properties = analyser.Properties(new Dictionary<string, object>
            {
                ["Title"] = "Project \"quoted\"", ["ListId"] = crossSite ? Guid.Empty : SourceId,
                ["PrimaryTaskListUrl"] = "Lists/Tasks", ["ShowTimeline"] = showTimeline
            }, WebParts.ProjectSummary,
                $"<webParts><webPart xmlns='http://schemas.microsoft.com/WebPart/v3'><data><properties><property name='Panels'>{WebUtility.HtmlEncode(Panels)}</property></properties></data></webPart></webParts>");
            Assert.AreEqual(Panels, properties["Panels"]);
            var webPart = new WebPartEntity { Type = WebParts.ProjectSummary, Properties = properties };
            var observer = new Mock<ILogObserver>();
            var processor = new FunctionProcessor(source, target, new Mock<PnPCore.IPage>().Object, transformation,
                new PageTransformationInformation(null) { SkipTelemetry = true }, new[] { observer.Object });
            processor.Process(ref mapping, webPart);
            var selected = mapping.Mappings.Mapping.Single();
            processor.ProcessMappingFunctions(ref mapping, webPart, selected);
            var output = selected.ClientSideWebPart.Single();
            Assert.AreEqual("7b90afac-708b-4317-be62-f51a74067d7e", output.ControlId);
            var json = JObject.Parse(WebUtility.HtmlDecode(TokenParser.ReplaceTokens(output.JsonControlData, webPart)));
            var values = json["properties"];
            var expectedUrl = crossSite ? "/sites/target/Lists/Tasks" : "/sites/source/Lists/Tasks";
            Assert.AreEqual((crossSite ? TargetId : SourceId).ToString(), (string)values["listId"]);
            Assert.AreEqual(expectedUrl, (string)values["primaryTaskListUrl"]);
            Assert.AreEqual(expectedUrl, (string)values["panels"][0]["listUrls"][0]);
            Assert.AreEqual(7, (int)values["panels"][0]["daysToSearch"]);
            Assert.AreEqual("Upcoming \"seven\"", (string)values["panels"][0]["title"]);
            Assert.AreEqual(showTimeline, (bool)values["showTimeline"]);
            if (showTimeline)
            {
                CollectionAssert.AreEqual(new[] { crossSite ? 9 : 1 }, values["timelineItemIds"].Values<int>().ToArray());
                Assert.AreEqual("Selected", (string)values["timelineViewName"]);
            }
            else
            {
                Assert.IsNull(values["timelineItemIds"]);
            }
            observer.Verify(log => log.Warning(It.Is<LogEntry>(entry => entry.Message.Contains("require remigration"))),
                showTimeline ? Times.Once() : Times.Never());
            observer.Verify(log => log.Error(It.IsAny<LogEntry>()), Times.Never);
        }

        [TestMethod]
        public void Timeline_PreservesSelectedTasksMilestonesCalloutsAndCustomDates()
        {
            var properties = new JObject();
            BuiltIn.AddProjectSummaryTimeline(properties, new Dictionary<string, object>
            {
                ["TimelineSelectedView"] = "Selected", ["TimelineDefaultView"] = "Other",
                ["Timeline_Other"] = "<TLViewData><t uid='99' onTL='1'/></TLViewData>",
                ["Timeline_Selected"] = "<TLViewData><t uid='2' onTL='1'/><m uid='3' onTL='1'/><ft uid='4' onTL='1'/><t uid='2' onTL='1'/><t uid='5' onTL='0'/><tl useCustomDates='1' startDate='2026-10-01' finishDate='2026-11-01'/></TLViewData>"
            });
            CollectionAssert.AreEqual(new[] { 2, 3, 4 }, properties["timelineItemIds"].Values<int>().ToArray());
            Assert.AreEqual("2026-10-01T00:00:00", (string)properties["timelineStartDate"]);
            Assert.AreEqual("2026-11-01T00:00:00", (string)properties["timelineEndDate"]);
        }

        [DataTestMethod]
        [DataRow(false)]
        [DataRow(true)]
        public void Timeline_EmptySelectionRemainsExplicitRatherThanEnablingHeuristic(bool configured)
        {
            var properties = new JObject();
            var folder = new Dictionary<string, object>();
            if (configured)
            {
                folder["Timeline_Default"] = "<TLViewData><t uid='2' onTL='0'/></TLViewData>";
            }
            BuiltIn.AddProjectSummaryTimeline(properties, folder);
            Assert.AreEqual(0, ((JArray)properties["timelineItemIds"]).Count);
        }

        [TestMethod]
        public void Timeline_RejectsAmbiguousSelectionAndUnsupportedMultipleBars()
        {
            Assert.ThrowsException<NotAvailableAtTargetException>(() => BuiltIn.AddProjectSummaryTimeline(new JObject(),
                new Dictionary<string, object> { ["Timeline_A"] = "<TLViewData/>", ["Timeline_B"] = "<TLViewData/>" }));
            Assert.ThrowsException<NotAvailableAtTargetException>(() => BuiltIn.AddProjectSummaryTimeline(new JObject(),
                new Dictionary<string, object> { ["Timeline_A"] = "<TLViewData><tl uid='1'/><tl uid='2'/></TLViewData>" }));
            Assert.ThrowsException<FormatException>(() => BuiltIn.AddProjectSummaryTimeline(new JObject(),
                new Dictionary<string, object> { ["Timeline_A"] = "<TLViewData><t uid='invalid' onTL='1'/></TLViewData>" }));
        }

        [TestMethod]
        public void Panels_EmptyConfigurationDoesNotEnableNativeDefaultsAndMissingExportFails()
        {
            using var source = CreateContext(SourceUrl, SourceId, 1);
            var builtIn = new BuiltIn(new PageTransformationInformation(null) { SkipTelemetry = true }, source, source);
            var properties = JObject.Parse(builtIn.ProjectSummaryProperties(SourceId, "", false, "[]", ""))["properties"];
            Assert.IsTrue(properties["panels"].All(panel => !(bool)panel["enabled"]));
            Assert.ThrowsException<NotAvailableAtTargetException>(() => builtIn.ProjectSummaryProperties(SourceId, "", false, "", ""));
        }

        private static ClientContext CreateContext(string url, Guid listId, int timelineId)
        {
            CacheManager.Instance.SetSharePointVersion(new Uri(url), SPVersion.SPO);
            var provider = new MockEntryResponseProvider();
            var web = new Dictionary<string, object>
            {
                ["_ObjectType_"] = "SP.Web", ["Url"] = url, ["ServerRelativeUrl"] = new Uri(url).AbsolutePath
            };
            foreach (var name in new[] { "Web", "RootWeb" })
            {
                provider.ResponseEntries.Add(new MockResponseEntry<object> { Url = url, PropertyName = name, ReturnValue = web });
            }
            provider.ResponseEntries.Add(new MockResponseEntry<object>
            {
                Url = url, PropertyName = "Site", ReturnValue = new Dictionary<string, object>
                {
                    ["_ObjectType_"] = "SP.Site", ["Url"] = url, ["RootWeb"] = web
                }
            });
            var list = new Dictionary<string, object>
            {
                ["_ObjectType_"] = "SP.List", ["Id"] = $"/Guid({listId})/", ["Title"] = "Tasks", ["BaseTemplate"] = 171,
                ["RootFolder"] = new Dictionary<string, object>
                {
                    ["_ObjectType_"] = "SP.Folder", ["ServerRelativeUrl"] = new Uri(url).AbsolutePath + "/Lists/Tasks",
                    ["Properties"] = new Dictionary<string, object>
                    {
                        ["_ObjectType_"] = "SP.PropertyValues", ["TimelineSelectedView"] = "Selected",
                        ["Timeline_Selected"] = $"<TLViewData><t uid='{timelineId}' onTL='1'/></TLViewData>"
                    }
                }
            };
            foreach (var method in new[] { "GetList", "GetListByTitle" })
            {
                provider.ResponseEntries.Add(new MockResponseEntry<object> { Url = url, Method = method, ReturnValue = list });
            }
            provider.ResponseEntries.Add(new MockResponseEntry<object>
            {
                Url = url, PropertyName = "Lists", ReturnValue = new Dictionary<string, object>
                {
                    ["_ObjectType_"] = "SP.ListCollection", ["_Child_Items_"] = new[] { list }
                }
            });
            return new ClientContext(url) { WebRequestExecutorFactory = new MockWebRequestExecutorFactory(provider) };
        }
    }
}
