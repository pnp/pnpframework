using AngleSharp.Html.Parser;
using Microsoft.SharePoint.Client;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Moq;
using Newtonsoft.Json.Linq;
using PnP.Framework.Modernization.Cache;
using PnP.Framework.Modernization.Entities;
using PnP.Framework.Modernization.Telemetry;
using PnP.Framework.Modernization.Transform;
using PnP.Framework.Utilities.UnitTests.Model;
using PnP.Framework.Utilities.UnitTests.Web;
using System;
using System.Collections.Generic;
using System.Linq;
using PnPCore = PnP.Core.Model.SharePoint;

namespace PnP.Framework.Modernization.Test.Transform.Functions
{
    [TestClass]
    public class BuiltInSiteFeedTests
    {
        private const string SourceUrl = "https://contoso.sharepoint.com/sites/source";
        private const string ComponentId = "b1d207ea-942e-4d3e-b890-a5c0fd61c670";

        [DataTestMethod]
        [DataRow(false)]
        [DataRow(true)]
        public void SameWeb_UsesNativeSiteFeedAndPreservesTitle(bool bracedComponentId)
        {
            const string title = "Feed \"quoted\" & <tag> \\ &quot;";
            using (var context = CreateContext(SourceUrl))
            {
                var observer = new Mock<ILogObserver>();
                var page = CreatePage(out var controls, bracedComponentId);

                Transform(context, context, page, observer, Part(WebParts.SiteFeed, title));

                var part = controls.Single() as PnPCore.IPageWebPart;
                Assert.IsNotNull(part);
                Assert.AreEqual(ComponentId, part.WebPartId);
                var json = JObject.Parse(part.PropertiesJson);
                Assert.AreEqual("1.0", (string)json["dataVersion"]);
                Assert.AreEqual(title, (string)json["properties"]["title"]);
                Assert.AreEqual(10, (int)json["properties"]["itemsPerPage"]);
                Assert.IsTrue((bool)json["properties"]["showReplies"]);
                observer.Verify(log => log.Warning(It.IsAny<LogEntry>()), Times.Never);
            }
        }

        [DataTestMethod]
        [DataRow(SourceUrl + "/subweb")]
        [DataRow("https://contoso.sharepoint.com/sites/target")]
        public void DifferentWeb_WarnsAndAddsSafePlaceholderWithoutStoppingOtherParts(string targetUrl)
        {
            using (var source = CreateContext(SourceUrl))
            using (var target = CreateContext(targetUrl))
            {
                const string title = "Feed <img src='x' onerror='alert(1)'> & \"quoted\"";
                var observer = new Mock<ILogObserver>();
                var page = CreatePage(out var controls);
                var following = Part("Example.UnmappedWebPart", "Following web part");
                following.Order = 2;

                Transform(source, target, page, observer, Part(WebParts.SiteFeed, title), following);

                Assert.AreEqual(2, controls.Count);
                using (var document = new HtmlParser().ParseDocument(((PnPCore.IPageText)controls[0]).Text))
                {
                    StringAssert.Contains(document.Body.TextContent, title);
                    StringAssert.Contains(document.Body.TextContent, "posts remain in the source site");
                    Assert.AreEqual(0, document.QuerySelectorAll("img, script").Length);
                }
                StringAssert.Contains(((PnPCore.IPageText)controls[1]).Text, "Following web part");
                observer.Verify(log => log.Warning(It.Is<LogEntry>(entry =>
                    entry.Message == string.Format(LogStrings.Warning_SiteFeedCrossSite, title, SourceUrl))), Times.Once);
                observer.Verify(log => log.Warning(It.IsAny<LogEntry>()), Times.Once);
                page.Verify(value => value.NewWebPart(It.IsAny<PnPCore.IPageComponent>()), Times.Never);
            }
        }

        private static void Transform(ClientContext source, ClientContext target, Mock<PnPCore.IPage> page,
            Mock<ILogObserver> observer, params WebPartEntity[] parts)
        {
            var information = new PageTransformationInformation(null) { SkipTelemetry = true };
            new ContentTransformator(source, target, page.Object, BasePageTransformator.LoadDefaultWebPartMapping(),
                information, new[] { observer.Object }).Transform(parts.ToList());
            observer.Verify(log => log.Error(It.IsAny<LogEntry>()), Times.Never);
        }

        private static WebPartEntity Part(string type, string title)
        {
            return new WebPartEntity
            {
                Type = type, Title = title, Row = 1, Column = 1, Order = 1,
                Properties = new Dictionary<string, string> { ["Title"] = title }
            };
        }

        private static Mock<PnPCore.IPage> CreatePage(out List<PnPCore.ICanvasControl> controls, bool bracedComponentId = false)
        {
            var added = new List<PnPCore.ICanvasControl>();
            controls = added;
            var page = new Mock<PnPCore.IPage>();
            page.SetupGet(value => value.PageTitle).Returns("Site Feed test");
            page.SetupGet(value => value.Controls).Returns(added);
            var column = new Mock<PnPCore.ICanvasColumn>();
            column.SetupGet(value => value.Controls).Returns(added);
            var section = new Mock<PnPCore.ICanvasSection>();
            section.SetupGet(value => value.Columns).Returns(new List<PnPCore.ICanvasColumn> { column.Object });
            page.SetupGet(value => value.Sections).Returns(new List<PnPCore.ICanvasSection> { section.Object });
            page.Setup(value => value.AvailablePageComponents(null)).Returns(new List<PnPCore.IPageComponent>
            {
                new PnPCore.PageComponent { Id = bracedComponentId ? "{" + ComponentId + "}" : ComponentId, Name = ComponentId }
            });
            page.Setup(value => value.NewWebPart(It.IsAny<PnPCore.IPageComponent>())).Returns<PnPCore.IPageComponent>(component =>
            {
                var part = new Mock<PnPCore.IPageWebPart>();
                part.SetupAllProperties();
                part.SetupGet(value => value.WebPartId).Returns(component.Id.Trim('{', '}'));
                return part.Object;
            });
            page.Setup(value => value.NewTextPart(It.IsAny<string>())).Returns(() =>
            {
                var text = new Mock<PnPCore.IPageText>();
                text.SetupAllProperties();
                return text.Object;
            });
            page.Setup(value => value.AddControl(It.IsAny<PnPCore.ICanvasControl>(), column.Object, It.IsAny<int>()))
                .Callback<PnPCore.ICanvasControl, PnPCore.ICanvasColumn, int, PnPCore.ControlFlexLayoutPosition>(
                    (control, _, order, position) => { control.Order = order; added.Add(control); });
            return page;
        }

        private static ClientContext CreateContext(string url)
        {
            CacheManager.Instance.SetSharePointVersion(new Uri(url), SPVersion.SPO);
            var provider = new MockEntryResponseProvider();
            var web = new Dictionary<string, object>
            {
                ["_ObjectType_"] = "SP.Web", ["Url"] = url, ["ServerRelativeUrl"] = new Uri(url).AbsolutePath,
                ["Id"] = $"/Guid({Guid.NewGuid()})/"
            };
            provider.ResponseEntries.Add(new MockResponseEntry<object> { Url = url, PropertyName = "Web", ReturnValue = web });
            provider.ResponseEntries.Add(new MockResponseEntry<object>
            {
                Url = url, PropertyName = "Site",
                ReturnValue = new Dictionary<string, object>
                {
                    ["_ObjectType_"] = "SP.Site", ["Url"] = url, ["ServerRelativeUrl"] = new Uri(url).AbsolutePath,
                    ["Id"] = $"/Guid({Guid.NewGuid()})/", ["RootWeb"] = web
                }
            });
            var context = new ClientContext(url) { WebRequestExecutorFactory = new MockWebRequestExecutorFactory(provider) };
            context.Load(context.Web, value => value.Id, value => value.Url, value => value.ServerRelativeUrl);
            context.Load(context.Site, value => value.Id, value => value.RootWeb.ServerRelativeUrl);
            context.ExecuteQuery();
            return context;
        }
    }
}
