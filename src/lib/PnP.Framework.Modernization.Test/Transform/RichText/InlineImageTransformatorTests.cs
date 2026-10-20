using AngleSharp.Html.Parser;
using Microsoft.SharePoint.Client;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Moq;
using PnP.Framework.Modernization;
using PnP.Framework.Modernization.Cache;
using PnP.Framework.Modernization.Entities;
using PnP.Framework.Modernization.Functions;
using PnP.Framework.Modernization.Telemetry;
using PnP.Framework.Modernization.Transform;
using PnP.Framework.Utilities.UnitTests.Model;
using PnP.Framework.Utilities.UnitTests.Web;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml.Serialization;
using PnPCore = PnP.Core.Model.SharePoint;

namespace PnP.Framework.Modernization.Test.Transform.RichText
{
    [TestClass]
    public class InlineImageTransformatorTests
    {
        private const string SiteUrl = "https://contoso.sharepoint.com/sites/source";
        private const string ImageUrl = "/sites/source/SiteAssets/image.png";
        private const string TargetImageUrl = "/sites/target/SiteAssets/SitePages/page/image.png";
        private static readonly Uri SourcePageUrl = new Uri(SiteUrl + "/Pages/page.aspx");

        [DataTestMethod]
        [DataRow("<p>before<img src='" + ImageUrl + "'>after</p>", "body > .imagePlugin")]
        [DataRow("<table><tbody><tr><td><p>before<img src='" + ImageUrl + "'>after</p></td></tr></tbody></table>", "td > .imagePlugin")]
        [DataRow("<ul><li><p>before<strong><img src='" + ImageUrl + "'></strong>after</p></li></ul>", "li > .imagePlugin")]
        [DataRow("<ol><li><ul><li><img src='" + ImageUrl + "'></li></ul></li></ol>", "ol > li > ul > li > .imagePlugin")]
        public void Transform_ConvertsInPlaceAndPreservesContainer(string html, string selector)
        {
            var page = CreatePage();
            var converted = CreateTransformator(page).Transform(html);

            using (var document = new HtmlParser().ParseDocument(converted))
            {
                Assert.AreEqual(1, document.QuerySelectorAll(selector).Length);
                Assert.AreEqual(0, document.QuerySelectorAll("img").Length);
                Assert.AreEqual(TargetImageUrl, document.QuerySelector(".imagePlugin").GetAttribute("data-imageurl"));
                Assert.AreEqual(0, document.QuerySelectorAll("p .imagePlugin, strong .imagePlugin").Length);
            }
            page.Verify(p => p.GetInlineImage(null, TargetImageUrl, It.IsAny<PnPCore.PageImageOptions>()), Times.Once);
            page.Verify(p => p.NewWebPart(It.IsAny<PnPCore.IPageComponent>()), Times.Never);
        }

        [TestMethod]
        public void Transform_PreservesLinkedSiblingTextAndImageOrder()
        {
            var html = $"<p id='original'>before<a href='/sites/target/Pages/details.aspx'>left<strong><img src='{ImageUrl}'></strong>middle<img src='{ImageUrl}'>right</a>after</p>";
            using (var document = new HtmlParser().ParseDocument(CreateTransformator(CreatePage()).Transform(html)))
            {
                var images = document.QuerySelectorAll(".imagePlugin");
                Assert.AreEqual(2, images.Length);
                Assert.IsTrue(images.All(image => image.GetAttribute("data-linkurl") == "/sites/target/Pages/details.aspx"));
                Assert.AreEqual("beforeleftmiddlerightafter", document.Body.TextContent);
                CollectionAssert.AreEqual(new[] { "P", "DIV", "P", "DIV", "P" },
                    document.Body.Children.Select(element => element.TagName).ToArray());
                Assert.AreEqual("beforeleft", document.Body.Children[0].TextContent);
                Assert.AreEqual("middle", document.Body.Children[2].TextContent);
                Assert.AreEqual("rightafter", document.Body.Children[4].TextContent);
                Assert.AreEqual(1, document.QuerySelectorAll("#original").Length);
                Assert.IsTrue(document.QuerySelectorAll("a").All(anchor => anchor.GetAttribute("href") == "/sites/target/Pages/details.aspx"));
            }
        }

        [TestMethod]
        public void Transform_PreservesDimensionsAndEncodesAltTextAndLinks()
        {
            PnPCore.PageImageOptions options = null;
            var page = CreatePage();
            page.Setup(p => p.GetInlineImage(null, TargetImageUrl, It.IsAny<PnPCore.PageImageOptions>()))
                .Callback<PnPCore.IPageText, string, PnPCore.PageImageOptions>((_, __, value) => options = value)
                .Returns<PnPCore.IPageText, string, PnPCore.PageImageOptions>((_, url, value) => ImageSnippet(url, value));
            var html = $"<a href='https://example.org/?a=1&amp;b=&quot;two&quot;'><img src='{ImageUrl}' alt='A &quot;quote&quot; &amp; &lt;tag&gt;' width='100' height='50' style='width: 240px; height: 120px; float: right'></a>";
            using (var document = new HtmlParser().ParseDocument(CreateTransformator(page).Transform(html)))
            {
                Assert.AreEqual(240, options.Width);
                Assert.AreEqual(120, options.Height);
                Assert.AreEqual(PnPCore.PageImageAlignment.Right, options.Alignment);
                var image = document.QuerySelector(".imagePlugin");
                Assert.AreEqual("A \"quote\" & <tag>", image.GetAttribute("data-alttext"));
                Assert.AreEqual("https://example.org/?a=1&b=\"two\"", image.GetAttribute("data-linkurl"));
                Assert.AreEqual(1, document.Body.Children.Length);
            }
        }

        [TestMethod]
        public void Transform_NormalizesLinkAroundBlockContainerWithoutLosingText()
        {
            var html = $"<a href='/details'><div>before<span><img src='{ImageUrl}'></span>after</div></a>";
            using (var document = new HtmlParser().ParseDocument(CreateTransformator(CreatePage()).Transform(html)))
            {
                Assert.AreEqual(1, document.QuerySelectorAll("body > .imagePlugin").Length);
                Assert.AreEqual("/details", document.QuerySelector(".imagePlugin").GetAttribute("data-linkurl"));
                CollectionAssert.AreEqual(new[] { "before", "after" },
                    document.QuerySelectorAll("a[href='/details']").Select(anchor => anchor.TextContent).ToArray());
            }
        }

        [TestMethod]
        public void Transform_PreservesPercentageWidth()
        {
            var page = CreatePage();
            CreateTransformator(page).Transform($"<img src='{ImageUrl}' width='300' height='200' style='width:50%'>");
            page.Verify(p => p.GetInlineImage(null, TargetImageUrl,
                It.Is<PnPCore.PageImageOptions>(options => options.Width == 300 && options.Height == 200 && options.WidthPercentage == 50)), Times.Once);
        }

        [DataTestMethod]
        [DataRow("", null)]
        [DataRow("style='width:50%'", "50")]
        public void Transform_DoesNotStretchPixelImagesToSdkDefaultPercentage(string style, string expectedPercentage)
        {
            var page = CreatePage();
            page.Setup(p => p.GetInlineImage(null, TargetImageUrl, It.IsAny<PnPCore.PageImageOptions>()))
                .Returns<PnPCore.IPageText, string, PnPCore.PageImageOptions>((_, url, options) =>
                {
                    options.WidthPercentage ??= 100;
                    return ImageSnippet(url, options);
                });
            var result = CreateTransformator(page).Transform($"<img src='{ImageUrl}' width='160' height='80' {style}>");
            using (var document = new HtmlParser().ParseDocument(result))
            {
                var image = document.QuerySelector(".imagePlugin");
                Assert.AreEqual("160", image.GetAttribute("data-width"));
                Assert.AreEqual("80", image.GetAttribute("data-height"));
                Assert.AreEqual(expectedPercentage, image.GetAttribute("data-widthpercentage"));
            }
        }

        [DataTestMethod]
        [DataRow("", 320, 160)]
        [DataRow("width='160'", 160, 80)]
        [DataRow("height='80'", 160, 80)]
        public void Transform_CompletesDimensionsFromImageMetadata(string attributes, int width, int height)
        {
            var page = CreatePage();
            var transformator = new InlineImageTransformator(page.Object, SourcePageUrl, _ => TargetImageUrl, value => value,
                imageLookup: _ => new Dictionary<string, string> { ["ImageWidth"] = "320", ["ImageHeight"] = "160" });
            var result = transformator.Transform($"<img src='{ImageUrl}' {attributes}>");
            page.Verify(p => p.GetInlineImage(null, TargetImageUrl,
                It.Is<PnPCore.PageImageOptions>(options => options.Width == width && options.Height == height)), Times.Once);
            using (var document = new HtmlParser().ParseDocument(result))
            {
                Assert.IsFalse(document.QuerySelector(".imagePlugin").HasAttribute("data-widthpercentage"));
            }
        }

        [TestMethod]
        public void Transform_UnknownAspectRatioRetainsHeightOnlyImageAndWarns()
        {
            var page = CreatePage();
            var observer = new Mock<ILogObserver>();
            var transformator = new InlineImageTransformator(page.Object, SourcePageUrl, _ => TargetImageUrl, value => value,
                new[] { observer.Object }, _ => new Dictionary<string, string> { ["ImageWidth"] = "-1", ["ImageHeight"] = "-1" });
            var html = $"<img src='{ImageUrl}' height='80'>";
            Assert.AreEqual(html, transformator.Transform(html));
            observer.Verify(log => log.Warning(It.Is<LogEntry>(entry => entry.Message.Contains("specified height"))), Times.Once);
            page.Verify(p => p.GetInlineImage(It.IsAny<PnPCore.IPageText>(), It.IsAny<string>(), It.IsAny<PnPCore.PageImageOptions>()), Times.Never);
        }

        [TestMethod]
        public void Transform_PreservesInheritedAlignmentAndIsIdempotent()
        {
            var page = CreatePage();
            var transformator = CreateTransformator(page);
            var html = $"<p style='text-align:center'>before<code><span><img src='{ImageUrl}'></span></code>after</p>";
            var result = transformator.Transform(html);

            page.Verify(p => p.GetInlineImage(null, TargetImageUrl,
                It.Is<PnPCore.PageImageOptions>(options => options.Alignment == PnPCore.PageImageAlignment.Center)), Times.Once);
            Assert.AreEqual(result, transformator.Transform(result));
            using (var document = new HtmlParser().ParseDocument(result))
            {
                Assert.AreEqual(1, document.QuerySelectorAll("body > .imagePlugin").Length);
                Assert.AreEqual("beforeafter", document.Body.TextContent);
            }
        }

        [DataTestMethod]
        [DataRow("../SiteAssets/image.png")]
        [DataRow(ImageUrl)]
        [DataRow(SiteUrl + "/SiteAssets/image.png")]
        [DataRow("//contoso.sharepoint.com/sites/source/SiteAssets/image.png")]
        public void Transform_ResolvesSourceBeforeAssetTransfer(string source)
        {
            string transferredSource = null;
            var transformator = new InlineImageTransformator(CreatePage().Object, SourcePageUrl,
                value => { transferredSource = value; return TargetImageUrl; }, value => value);

            transformator.Transform($"<img src='{source}'>");

            Assert.AreEqual(ImageUrl, transferredSource);
        }

        [DataTestMethod]
        [DataRow("/sites/target/SiteAssets/image.png")]
        [DataRow("/sites/target/siteassets/IMAGE.PNG")]
        public void Transform_RewritesImageSelfLinkToCopiedAsset(string link)
        {
            var transformator = new InlineImageTransformator(CreatePage().Object, SourcePageUrl,
                _ => TargetImageUrl, value => value.Replace("/sites/source/", "/sites/target/"));
            var result = transformator.Transform($"<a href='{link}'><img src='{ImageUrl}'></a>");
            using (var document = new HtmlParser().ParseDocument(result))
            {
                Assert.AreEqual(TargetImageUrl, document.QuerySelector(".imagePlugin").GetAttribute("data-linkurl"));
            }
        }

        [DataTestMethod]
        [DataRow("https://external.example/image.png")]
        [DataRow("data:image/png;base64,AAAA")]
        [DataRow("")]
        public void Transform_UnsupportedImageRetainsOriginalAndWarns(string source)
        {
            var page = CreatePage();
            var observer = new Mock<ILogObserver>();
            var html = $"<p><a href='/details'><img src=\"{source}\" alt=\"keep\" style=\"width: 42px; height: 21px\"></a></p>";
            var result = CreateTransformator(page, observer.Object).Transform(html);

            Assert.AreEqual(html, result);
            observer.Verify(log => log.Warning(It.Is<LogEntry>(entry => entry.Message.Contains("Retaining the original image HTML"))), Times.Once);
            page.Verify(p => p.GetInlineImage(It.IsAny<PnPCore.IPageText>(), It.IsAny<string>(), It.IsAny<PnPCore.PageImageOptions>()), Times.Never);
        }

        [TestMethod]
        public void Transform_MissingImageRetainsOriginalButConvertsOtherImages()
        {
            var page = CreatePage();
            var observer = new Mock<ILogObserver>();
            page.SetupSequence(p => p.GetInlineImage(null, TargetImageUrl, It.IsAny<PnPCore.PageImageOptions>()))
                .Throws(new FileNotFoundException("Missing image"))
                .Returns(ImageSnippet(TargetImageUrl, new PnPCore.PageImageOptions()));

            var html = $"<p><img src='{ImageUrl}' alt='missing'><img src='{ImageUrl}' alt='present'></p>";
            using (var document = new HtmlParser().ParseDocument(CreateTransformator(page, observer.Object).Transform(html)))
            {
                Assert.AreEqual(ImageUrl, document.QuerySelector("p > img").GetAttribute("src"));
                Assert.AreEqual("missing", document.QuerySelector("img").GetAttribute("alt"));
                Assert.AreEqual("present", document.QuerySelector(".imagePlugin").GetAttribute("data-alttext"));
            }
            observer.Verify(log => log.Warning(It.Is<LogEntry>(entry => entry.Message.Contains("Missing image"))), Times.Once);
        }

        [TestMethod]
        public void Transform_CK4TargetFailsBeforeTransferringAssets()
        {
            var page = CreatePage();
            page.SetupGet(p => p.EditorType).Returns(PnPCore.EditorType.CK4);
            var observer = new Mock<ILogObserver>();
            var transformator = new InlineImageTransformator(page.Object, SourcePageUrl,
                _ => throw new AssertFailedException("No asset transfer should occur"), value => value, new[] { observer.Object });

            Assert.ThrowsException<NotSupportedException>(() => transformator.Transform($"<img src='{ImageUrl}'>"));
            observer.Verify(log => log.Error(It.Is<LogEntry>(entry => entry.Message.Contains("CK5"))), Times.Once);
        }

        [TestMethod]
        public void Transform_DoesNotHideUnexpectedProgrammingErrors()
        {
            var page = CreatePage();
            page.Setup(p => p.GetInlineImage(null, TargetImageUrl, It.IsAny<PnPCore.PageImageOptions>()))
                .Throws(new InvalidOperationException("Unexpected"));
            Assert.ThrowsException<InvalidOperationException>(() => CreateTransformator(page).Transform($"<img src='{ImageUrl}'>"));
        }

        [TestMethod]
        public void TextCleanup_PreservesImageSourceAndAttributesThroughUrlAndTableCleanup()
        {
            var html = $"<table><thead><tr><th><a href='/sites/source/Pages/details.aspx'><img src='{ImageUrl}' style='height: 42px; width: 84px' alt='A &amp; B'></a></th></tr></thead><tbody><tr><td><ul><li><img src='{ImageUrl}'></li></ul></td></tr></tbody></table>";
            var result = InlineImageTransformator.PreserveImages(html, text =>
                new HtmlTransformator().Transform(text.Replace("/sites/source/", "/sites/target/"), true));

            using (var document = new HtmlParser().ParseDocument(result))
            {
                Assert.AreEqual(2, document.QuerySelectorAll("img").Length);
                Assert.IsTrue(document.QuerySelectorAll("img").All(image => image.GetAttribute("src") == ImageUrl));
                Assert.AreEqual("height: 42px; width: 84px", document.QuerySelector("img").GetAttribute("style"));
                Assert.AreEqual("/sites/target/Pages/details.aspx", document.QuerySelector("a").GetAttribute("href"));
                Assert.AreEqual("A & B", document.QuerySelector("img").GetAttribute("alt"));
                Assert.IsFalse(result.Contains("placeholder"));
                Assert.IsFalse(result.Contains("pnpimage"));
            }

            using (var document = new HtmlParser().ParseDocument(CreateTransformator(CreatePage()).Transform(result)))
            {
                Assert.AreEqual(2, document.QuerySelectorAll("td .imagePlugin").Length);
                Assert.AreEqual(1, document.QuerySelectorAll("li > .imagePlugin").Length);
            }
        }

        [TestMethod]
        public void TextCleanup_PreservesImagesInAllTableSectionsAndHeaderRows()
        {
            var html = $"<table><thead><tr><th><img src='{ImageUrl}' alt='header1'></th></tr><tr><th><img src='{ImageUrl}' alt='header2'></th></tr></thead><tbody><tr><td><img src='{ImageUrl}' alt='body1'></td></tr></tbody><tbody><tr><td><img src='{ImageUrl}' alt='body2'></td></tr></tbody><tfoot><tr><td><img src='{ImageUrl}' alt='footer'></td></tr></tfoot></table>";
            var result = InlineImageTransformator.PreserveImages(html, text => new HtmlTransformator().Transform(text, false));
            using (var document = new HtmlParser().ParseDocument(CreateTransformator(CreatePage()).Transform(result)))
            {
                CollectionAssert.AreEqual(new[] { "header1", "header2", "body1", "body2", "footer" },
                    document.QuerySelectorAll("td .imagePlugin").Select(image => image.GetAttribute("data-alttext")).ToArray());
                Assert.AreEqual(5, document.QuerySelectorAll("tr").Length);
            }
        }

        [TestMethod]
        public void TextCleanup_WithoutImagesRetainsExistingOutput()
        {
            const string html = "<h1>Heading</h1><p>Text <a href='/sites/source/Pages/other.aspx'>link</a></p><iframe src='https://video.example'></iframe>";
            Func<string, string> cleanup = value => new HtmlTransformator().Transform(value.Replace("/sites/source/", "/sites/target/"), true);
            Assert.AreEqual(cleanup(html), InlineImageTransformator.PreserveImages(html, cleanup));
        }

        [TestMethod]
        public void Transform_AssetTransferFailureRetainsSourceAndWarns()
        {
            var observer = new Mock<ILogObserver>();
            var page = CreatePage();
            var transformator = new InlineImageTransformator(page.Object, SourcePageUrl,
                _ => throw new UnauthorizedAccessException("Cannot copy asset"), value => value, new[] { observer.Object });
            var html = $"<p><img src='{ImageUrl}' alt='keep'></p>";

            Assert.AreEqual(html, transformator.Transform(html));
            page.Verify(p => p.GetInlineImage(It.IsAny<PnPCore.IPageText>(), It.IsAny<string>(), It.IsAny<PnPCore.PageImageOptions>()), Times.Never);
            observer.Verify(log => log.Warning(It.Is<LogEntry>(entry => entry.Message.Contains("Cannot copy asset"))), Times.Once);
        }

        [DataTestMethod]
        [DataRow(true, true)]
        [DataRow(true, false)]
        [DataRow(false, true)]
        [DataRow(false, false)]
        public void WikiSplitter_NeverSplitsOrExtractsImages(bool handleImagesAndVideos, bool addTableListImages)
        {
            using (var context = CreateContext())
            {
                var original = WikiText($"<p>before<img src='{ImageUrl}'>after</p><table><tr><td><img src='{ImageUrl}'></td></tr></table><ul><li><a href='/link'><img src='{ImageUrl}'></a></li></ul>");
                var result = new WikiHtmlTransformator(context, context, CreatePage().Object, new PageTransformationInformation(null))
                    .TransformPlusSplit(new List<WebPartEntity> { original }, handleImagesAndVideos, addTableListImages);

                Assert.AreEqual(1, result.Count);
                Assert.AreSame(original, result[0]);
                Assert.AreEqual(3, new HtmlParser().ParseDocument(result[0].Properties["Text"]).QuerySelectorAll("img").Length);
            }
        }

        [DataTestMethod]
        [DataRow(true)]
        [DataRow(false)]
        public void WikiSplitter_PreservesIframeBehaviorWithImages(bool splitVideo)
        {
            using (var context = CreateContext())
            {
                var result = new WikiHtmlTransformator(context, context, CreatePage().Object, new PageTransformationInformation(null))
                    .TransformPlusSplit(new List<WebPartEntity>
                    {
                        WikiText($"<p>before<img src='{ImageUrl}'></p><iframe src='https://video.example/embed'></iframe><p>after</p>")
                    }, splitVideo, true);

                Assert.AreEqual(1, result.Count(part => part.Type == WebParts.WikiVideo));
                Assert.AreEqual(0, result.Count(part => part.Type == WebParts.WikiImage));
                Assert.AreEqual(splitVideo ? 2 : 1, result.Count(part => part.Type == WebParts.WikiText));
                Assert.AreEqual(1, result.Where(part => part.Type == WebParts.WikiText)
                    .Sum(part => new HtmlParser().ParseDocument(part.Properties["Text"]).QuerySelectorAll("img").Length));
            }
        }

        [DataTestMethod]
        [DataRow(false)]
        [DataRow(true)]
        public void ContentMapping_ProducesOneTextControlWithNativeImages(bool contentEditor)
        {
            using (var context = CreateContext())
            {
                var page = CreatePage();
                var controls = ConfigureCanvas(page);
                var information = new PageTransformationInformation(null);
                var source = WikiText($"<p>before<img src='{ImageUrl}' alt='inline'>after</p>");
                if (contentEditor)
                {
                    source.Type = WebParts.ContentEditor;
                    source.Properties = new Dictionary<string, string>
                    {
                        ["Content"] = source.Properties["Text"], ["ContentLink"] = "", ["ChromeType"] = "2"
                    };
                }
                var parts = new WikiHtmlTransformator(context, context, page.Object, information)
                    .TransformPlusSplit(new List<WebPartEntity> { source }, true, true);
                new ContentTransformator(context, context, page.Object, LoadMapping(), information).Transform(parts);

                Assert.AreEqual(1, controls.Count);
                var text = controls.Single() as PnPCore.IPageText;
                Assert.IsNotNull(text);
                using (var document = new HtmlParser().ParseDocument(text.Text))
                {
                    Assert.AreEqual(1, document.QuerySelectorAll(".imagePlugin").Length);
                    Assert.AreEqual(0, document.QuerySelectorAll("img").Length);
                    Assert.AreEqual("beforeafter", document.Body.TextContent);
                    Assert.AreEqual(ImageUrl, document.QuerySelector(".imagePlugin").GetAttribute("data-imageurl"));
                }
                page.Verify(p => p.GetInlineImage(null, ImageUrl, It.IsAny<PnPCore.PageImageOptions>()), Times.Once);
            }
        }

        [TestMethod]
        public void ContentMapping_ScriptContentEditorDoesNotConvertOrTransferImages()
        {
            using (var context = CreateContext())
            {
                var page = CreatePage();
                var controls = ConfigureCanvas(page);
                var information = new PageTransformationInformation(null);
                var source = new WebPartEntity
                {
                    Type = WebParts.ContentEditor, Row = 1, Column = 1, Order = 1,
                    Properties = new Dictionary<string, string>
                    {
                        ["Content"] = $"<script>alert(1)</script><img src='{ImageUrl}'>",
                        ["ContentLink"] = "", ["ChromeType"] = "2"
                    }
                };
                var parts = new WikiHtmlTransformator(context, context, page.Object, information)
                    .TransformPlusSplit(new List<WebPartEntity> { source }, true, true);
                new ContentTransformator(context, context, page.Object, LoadMapping(), information).Transform(parts);

                Assert.AreEqual("Content editor contains script and could not be loaded.", ((PnPCore.IPageText)controls.Single()).Text);
                page.Verify(p => p.GetInlineImage(It.IsAny<PnPCore.IPageText>(), It.IsAny<string>(), It.IsAny<PnPCore.PageImageOptions>()), Times.Never);
            }
        }

        [TestMethod]
        public void ContentMapping_MissingImagePreservesItsSourceAndStyleAndWarns()
        {
            using (var context = CreateContext())
            {
                var page = CreatePage();
                var controls = ConfigureCanvas(page);
                var observer = new Mock<ILogObserver>();
                page.Setup(p => p.GetInlineImage(null, ImageUrl, It.IsAny<PnPCore.PageImageOptions>()))
                    .Throws(new FileNotFoundException("Missing image"));
                var source = WikiText($"<p>before<img src='{ImageUrl}' alt='keep' style='height: 42px; width: 84px'>after</p>");
                new ContentTransformator(context, context, page.Object, LoadMapping(), new PageTransformationInformation(null),
                    new[] { observer.Object }).Transform(new List<WebPartEntity> { source });

                using (var document = new HtmlParser().ParseDocument(((PnPCore.IPageText)controls.Single()).Text))
                {
                    Assert.AreEqual(ImageUrl, document.QuerySelector("p > img").GetAttribute("src"));
                    Assert.AreEqual("height: 42px; width: 84px", document.QuerySelector("img").GetAttribute("style"));
                    Assert.AreEqual("beforeafter", document.Body.TextContent);
                    Assert.AreEqual(0, document.QuerySelectorAll(".imagePlugin").Length);
                }
                observer.Verify(log => log.Warning(It.Is<LogEntry>(entry => entry.Message.Contains("Missing image"))), Times.Once);
            }
        }

        [DataTestMethod]
        [DataRow(false)]
        [DataRow(true)]
        public void BuiltInCleanup_PreservesSourceImageWhileRespectingLinkRewriteOptions(bool skipUrlRewrite)
        {
            using (var source = CreateContext())
            using (var target = CreateContext("https://contoso.sharepoint.com/sites/target"))
            {
                var information = new PageTransformationInformation(null) { SkipUrlRewrite = skipUrlRewrite };
                var builtIn = new BuiltIn(information, target, source, CreatePage().Object);
                var result = builtIn.TextCleanup($"<a href='/sites/source/Pages/details.aspx'><img src='{ImageUrl}'></a>", "true");
                using (var document = new HtmlParser().ParseDocument(result))
                {
                    Assert.AreEqual(ImageUrl, document.QuerySelector("img").GetAttribute("src"));
                    Assert.AreEqual(skipUrlRewrite ? "/sites/source/Pages/details.aspx" : "/sites/target/Pages/details.aspx",
                        document.QuerySelector("a").GetAttribute("href"));
                    Assert.IsFalse(result.Contains("placeholder"));
                }
            }
        }

        [TestMethod]
        public void BuiltInConversion_ResolvesRelativeImagesAgainstSourcePageFileRef()
        {
            using (var context = CreateContext())
            {
                var sourcePage = context.Web.Lists.GetByTitle("Pages").GetItemById(1);
                sourcePage[Constants.FileRefField] = "/sites/source/Pages/folder/page.aspx";
                var page = CreatePage();
                var builtIn = new BuiltIn(new PageTransformationInformation(sourcePage), context, context, page.Object);

                builtIn.TransformInlineImages("<img src='../../SiteAssets/image.png' width='160' height='80'>");

                page.Verify(p => p.GetInlineImage(null, ImageUrl, It.IsAny<PnPCore.PageImageOptions>()), Times.Once);
            }
        }

        private static InlineImageTransformator CreateTransformator(Mock<PnPCore.IPage> page, ILogObserver observer = null)
        {
            return new InlineImageTransformator(page.Object, SourcePageUrl, _ => TargetImageUrl, value => value,
                observer == null ? null : new[] { observer });
        }

        private static Mock<PnPCore.IPage> CreatePage()
        {
            var page = new Mock<PnPCore.IPage>();
            page.SetupGet(p => p.EditorType).Returns(PnPCore.EditorType.CK5);
            page.SetupGet(p => p.PageTitle).Returns("page");
            page.Setup(p => p.GetInlineImage(null, It.IsAny<string>(), It.IsAny<PnPCore.PageImageOptions>()))
                .Returns<PnPCore.IPageText, string, PnPCore.PageImageOptions>((_, url, options) => ImageSnippet(url, options));
            return page;
        }

        private static string ImageSnippet(string url, PnPCore.PageImageOptions options)
        {
            return $"<div class=\"imagePlugin\" style=\"background-color:transparent;position:relative;\" data-alignment=\"{options.Alignment}\" data-imageurl=\"{url}\" data-uploading=\"0\" data-height=\"{options.Height.GetValueOrDefault()}\" data-width=\"{options.Width.GetValueOrDefault()}\" data-widthpercentage=\"{options.WidthPercentage ?? 100}\" data-alttext=\"\" data-linkurl=\"\"></div>";
        }

        private static List<PnPCore.ICanvasControl> ConfigureCanvas(Mock<PnPCore.IPage> page)
        {
            var controls = new List<PnPCore.ICanvasControl>();
            var column = new Mock<PnPCore.ICanvasColumn>();
            column.SetupGet(c => c.Controls).Returns(controls);
            var section = new Mock<PnPCore.ICanvasSection>();
            section.SetupGet(s => s.Columns).Returns(new List<PnPCore.ICanvasColumn> { column.Object });
            page.SetupGet(p => p.Sections).Returns(new List<PnPCore.ICanvasSection> { section.Object });
            page.Setup(p => p.AvailablePageComponents(null)).Returns(new List<PnPCore.IPageComponent>());
            page.Setup(p => p.NewTextPart(It.IsAny<string>())).Returns(() =>
            {
                var text = new Mock<PnPCore.IPageText>();
                text.SetupAllProperties();
                return text.Object;
            });
            page.Setup(p => p.AddControl(It.IsAny<PnPCore.ICanvasControl>(), column.Object, It.IsAny<int>()))
                .Callback<PnPCore.ICanvasControl, PnPCore.ICanvasColumn, int, PnPCore.ControlFlexLayoutPosition>((control, _, order, position) => controls.Add(control));
            return controls;
        }

        private static WebPartEntity WikiText(string html)
        {
            return new WebPartEntity
            {
                Type = WebParts.WikiText, Row = 1, Column = 1, Order = 1,
                Properties = new Dictionary<string, string> { ["Text"] = html }
            };
        }

        private static PageTransformation LoadMapping()
        {
            using (var reader = new StringReader(BasePageTransformator.LoadDefaultWebPartMappingFile()))
            {
                return (PageTransformation)new XmlSerializer(typeof(PageTransformation)).Deserialize(reader);
            }
        }

        private static ClientContext CreateContext(string siteUrl = SiteUrl)
        {
            CacheManager.Instance.SetSharePointVersion(new Uri(siteUrl), SPVersion.SPO);
            var provider = new MockEntryResponseProvider();
            var web = new Dictionary<string, object>
            {
                ["_ObjectType_"] = "SP.Web", ["Url"] = siteUrl, ["ServerRelativeUrl"] = new Uri(siteUrl).AbsolutePath,
                ["Id"] = "/Guid(413af984-d496-45ef-82d6-a2835dca7d49)/"
            };
            provider.ResponseEntries.Add(new MockResponseEntry<object>
            {
                Url = siteUrl, PropertyName = "Web", ReturnValue = web
            });
            provider.ResponseEntries.Add(new MockResponseEntry<object>
            {
                Url = siteUrl, PropertyName = "Site",
                ReturnValue = new Dictionary<string, object>
                {
                    ["_ObjectType_"] = "SP.Site", ["Url"] = siteUrl, ["ServerRelativeUrl"] = new Uri(siteUrl).AbsolutePath,
                    ["Id"] = "/Guid(e7f8587b-0764-4b08-a5dd-bd33683daa4e)/", ["RootWeb"] = web
                }
            });
            provider.ResponseEntries.Add(new MockResponseEntry<object>
            {
                Url = siteUrl, Method = "GetFileByServerRelativeUrl",
                ReturnValue = new Dictionary<string, object>
                {
                    ["_ObjectType_"] = "SP.File",
                    ["ListId"] = "/Guid(80e43d5d-5fd0-48f3-86ea-db933db5e21b)/",
                    ["UniqueId"] = "/Guid(296c36c4-8b73-4fc7-9800-64855534d108)/",
                    ["Properties"] = new Dictionary<string, object>
                    {
                        ["_ObjectType_"] = "SP.PropertyValues", ["vti_lastwidth"] = "320", ["vti_lastheight"] = "160"
                    }
                }
            });
            var context = new ClientContext(siteUrl)
            {
                WebRequestExecutorFactory = new MockWebRequestExecutorFactory(provider)
            };
            context.Load(context.Web, value => value.Id, value => value.Url, value => value.ServerRelativeUrl);
            context.Load(context.Site, value => value.Id, value => value.Url, value => value.ServerRelativeUrl,
                value => value.RootWeb.ServerRelativeUrl);
            context.ExecuteQuery();
            return context;
        }
    }
}
