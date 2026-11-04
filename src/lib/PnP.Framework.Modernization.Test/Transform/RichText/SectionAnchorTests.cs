using AngleSharp.Dom;
using AngleSharp.Html.Dom;
using AngleSharp.Html.Parser;
using Microsoft.SharePoint.Client;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using PnP.Framework.Modernization.Functions;
using PnP.Framework.Modernization.Transform;
using PnP.Framework.Utilities.UnitTests.Model;
using PnP.Framework.Utilities.UnitTests.Web;
using System;
using System.Collections.Generic;
using System.Linq;

namespace PnP.Framework.Modernization.Test.Transform.RichText
{
    [TestClass]
    public class SectionAnchorTests
    {
        [DataTestMethod]
        [DataRow("h1", "h2")]
        [DataRow("h2", "h3")]
        [DataRow("h3", "h4")]
        [DataRow("h4", "div")]
        [DataRow("h5", "div")]
        [DataRow("h6", "div")]
        public void Transform_PreservesHeadingTargets(string sourceTag, string targetTag)
        {
            var html = $"<a href='#Service'>Service</a><{sourceTag} id='Service' " +
                "class='ms-rteForeColor-6' style='text-align:center'><span id='nested'>Section</span>" +
                $"</{sourceTag}>";
            using (var result = TransformAndAssertTargets(html))
            {
                var target = result.GetElementById("Service");
                Assert.AreEqual(targetTag.ToUpperInvariant(),
                    targetTag == "div" ? target.TagName : target.ParentElement.TagName);
                if (targetTag != "div")
                {
                    Assert.AreEqual("SPAN", target.TagName);
                    Assert.IsFalse(target.ParentElement.HasAttribute("id"),
                        "Modern rendering replaces heading IDs, so only the content must carry the legacy ID.");
                    StringAssert.Contains(target.ParentElement.GetAttribute("style"), "text-align: center");
                }
                Assert.AreEqual("Section", result.GetElementById("nested").TextContent);
                Assert.IsNotNull(result.QuerySelector(".fontColorGreen"));
            }
        }

        [DataTestMethod]
        [DataRow("span", "", "")]
        [DataRow("span", "", "Section")]
        [DataRow("span", "class='ms-rteStyle-Normal ms-rteFontFace-1'", "Section")]
        [DataRow("span", "class='ms-rteStyle-Quote'", "Section")]
        [DataRow("span", "class='ms-rteStyle-IntenseEmphasis'", "Section")]
        [DataRow("span", "class='ms-rteForeColor-6' style='text-decoration:underline'", "Section")]
        [DataRow("span", "style='text-decoration:line-through'", "Section")]
        [DataRow("strong", "style='text-decoration:underline'", "Section")]
        [DataRow("em", "style='text-decoration:line-through'", "Section")]
        [DataRow("sup", "style='text-decoration:underline'", "Section")]
        [DataRow("sub", "style='text-decoration:line-through'", "Section")]
        [DataRow("font", "style='text-decoration:underline'", "Section")]
        [DataRow("p", "style='text-decoration:line-through'", "Section")]
        public void Transform_PreservesReplacedAndUnwrappedTargets(string tag, string attributes, string content)
        {
            var html = $"<a href='#third'>Section</a><{tag} id='third' {attributes}>{content}</{tag}>";
            using (var result = TransformAndAssertTargets(html))
            {
                Assert.AreEqual(content, result.GetElementById("third").TextContent);
                using (var withoutId = new HtmlParser().ParseDocument(
                    new HtmlTransformator().Transform(html.Replace(" id='third'", ""), false)))
                {
                    CollectionAssert.AreEqual(
                        withoutId.QuerySelectorAll("em,u,s,strong,sup,sub,.fontColorGreen")
                            .Select(element => element.TagName + ":" + element.TextContent).ToArray(),
                        result.QuerySelectorAll("em,u,s,strong,sup,sub,.fontColorGreen")
                            .Select(element => element.TagName + ":" + element.TextContent).ToArray());
                }
            }
        }

        [TestMethod]
        public void Transform_PreservesNestedTargetsWithoutDuplicatingIdsOnWrappers()
        {
            const string html = "<a href='#outer'>Outer</a><a href='#inner'>Inner</a>" +
                "<span id='outer' class='ms-rteStyle-IntenseEmphasis'><span id='inner'>Section</span></span>" +
                "<p id='paragraph'><strong id='bold'>Bold</strong><font id='font' color='green'>Green</font></p>";
            using (var result = TransformAndAssertTargets(html))
            {
                Assert.IsTrue(result.GetElementById("outer").Contains(result.GetElementById("inner")));
            }
        }

        [DataTestMethod]
        [DataRow("false")]
        [DataRow("true")]
        public void TextCleanup_PreservesReportedSectionLinksAndTargets(string usePlaceHolders)
        {
            var headingIds = new[] { "overview", "Service", "ESRPAuth", "Integration", "summary", "onboarding" };
            var spanIds = new[] { "third", "azure", "Register" };
            var html = string.Concat(headingIds.Concat(spanIds).Select(id => $"<a href='#{id}'>{id}</a>")) +
                "<a href='#external%20'>Already broken</a>" +
                string.Concat(headingIds.Select(id => $"<h2 id='{id}'>{id}</h2>")) +
                "<span id='third'>Third</span><span id='azure' class='ms-rteStyle-Normal'>Azure</span>" +
                "<span id='Register'><strong>Register</strong></span>";
            const string siteUrl = "https://contoso.sharepoint.com/sites/wiki";
            var provider = new MockEntryResponseProvider();
            foreach (var property in new[] { "Site", "Web" })
            {
                provider.ResponseEntries.Add(new MockResponseEntry<object>
                {
                    Url = siteUrl,
                    PropertyName = property,
                    ReturnValue = new Dictionary<string, object>
                    {
                        ["_ObjectType_"] = "SP." + property,
                        ["Url"] = siteUrl,
                        ["ServerRelativeUrl"] = new Uri(siteUrl).AbsolutePath
                    }
                });
            }
            using (var context = new ClientContext(siteUrl)
            {
                WebRequestExecutorFactory = new MockWebRequestExecutorFactory(provider)
            })
            {
                var builtIn = new BuiltIn(new PageTransformationInformation(null) { SkipUrlRewrite = true }, context);
                using (var source = new HtmlParser().ParseDocument(html))
                using (var result = new HtmlParser().ParseDocument(builtIn.TextCleanup(html, usePlaceHolders)))
                {
                    AssertTargets(source, result);
                    Assert.IsNull(result.GetElementById("external "));
                    Assert.AreEqual("#external%20", result.QuerySelectorAll("a").Last().GetAttribute("href"));
                }
            }
        }

        [TestMethod]
        public void Transform_StillUnwrapsPlainSpansWithoutTargets()
        {
            Assert.AreEqual("Text", new HtmlTransformator().Transform("<span>Text</span><span></span>", false));
        }

        [DataTestMethod]
        [DataRow("overview", "Overview")]
        [DataRow("Section:1.2", "Section")]
        [DataRow("space here", "Section")]
        [DataRow("empty", "")]
        public void Transform_PreservesUnstyledHeadingAnchorInsideHeading(string id, string content)
        {
            var html = $"<a href='#{Uri.EscapeDataString(id)}'>Section</a><h2 id='{id}'>{content}</h2>";
            using (var result = TransformAndAssertTargets(html))
            {
                Assert.AreEqual("SPAN", result.GetElementById(id).TagName);
                Assert.AreEqual("H3", result.GetElementById(id).ParentElement.TagName);
            }
        }

        private static IHtmlDocument TransformAndAssertTargets(string html)
        {
            var result = new HtmlParser().ParseDocument(new HtmlTransformator().Transform(html, false));
            using (var source = new HtmlParser().ParseDocument(html))
            {
                AssertTargets(source, result);
            }
            return result;
        }

        private static void AssertTargets(IHtmlDocument source, IHtmlDocument result)
        {
            Assert.AreEqual(source.Body.TextContent, result.Body.TextContent);
            CollectionAssert.AreEquivalent(
                source.QuerySelectorAll("[id]").Select(element => element.Id).ToArray(),
                result.QuerySelectorAll("[id]").Select(element => element.Id).ToArray());
            CollectionAssert.AreEqual(
                source.QuerySelectorAll("a[href]").Select(element => element.GetAttribute("href")).ToArray(),
                result.QuerySelectorAll("a[href]").Select(element => element.GetAttribute("href")).ToArray());
            foreach (var link in source.QuerySelectorAll("a[href^='#']"))
            {
                var id = Uri.UnescapeDataString(link.GetAttribute("href").Substring(1));
                var target = source.GetElementById(id);
                if (target != null)
                {
                    Assert.IsNotNull(result.GetElementById(id), $"Missing destination for #{id}");
                    Assert.AreEqual(target.TextContent, result.GetElementById(id).TextContent);
                }
            }
        }
    }
}
