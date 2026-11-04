using AngleSharp.Html.Parser;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Moq;
using PnP.Framework.Modernization.Telemetry;
using PnP.Framework.Modernization.Transform;
using System.Linq;

namespace PnP.Framework.Modernization.Tests.Transform.Tables
{
    [TestClass]
    public class HtmlTableTransformatorTests
    {
        [DataTestMethod]
        [DataRow(0)]
        [DataRow(1)]
        [DataRow(2)]
        public void Transform_UnsizedSingleRowBorderlessTableUsesAutomaticWidths(int contentColumn)
        {
            var cells = Enumerable.Range(0, 3).Select(column => column == contentColumn
                ? "<td><strong>Troubleshooting instructions:</strong><p>Read the guidance.</p></td>"
                : "<td>\u200b<br></td>");
            var html = $"<table><tr>{string.Concat(cells)}</tr></table>";

            using (var document = new HtmlParser().ParseDocument(new HtmlTransformator().Transform(html, false)))
            {
                var result = document.QuerySelectorAll("table tr td");
                Assert.AreEqual(3, result.Length);
                for (int i = 0; i < result.Length; i++)
                {
                    Assert.AreEqual("width: auto", result[i].GetAttribute("style").TrimEnd(';'));
                }
                Assert.AreEqual("Troubleshooting instructions:", result[contentColumn].QuerySelector("strong").TextContent);
                Assert.AreEqual("Read the guidance.", result[contentColumn].QuerySelector("p").TextContent);
            }
        }

        [DataTestMethod]
        [DataRow("<table><tr><td>First</td><td>Second</td><td></td></tr></table>")]
        [DataRow("<table><tr><td>First</td><td><img src='/image.png'></td><td></td></tr></table>")]
        [DataRow("<table><tr><td>First</td><td><hr></td><td></td></tr></table>")]
        [DataRow("<table><tr><td>First</td><td><svg></svg></td><td></td></tr></table>")]
        [DataRow("<table><tr><td>First</td><td style='background-color:red'></td><td></td></tr></table>")]
        [DataRow("<table><tr><td>First</td><td class='custom-cell'></td><td></td></tr></table>")]
        [DataRow("<table><tr><td>First</td><td><span style='width:25px'></span></td><td></td></tr></table>")]
        [DataRow("<table border='1'><tr><td>First</td><td></td><td></td></tr></table>")]
        [DataRow("<table border='0' class='custom'><tr><td>First</td><td></td><td></td></tr></table>")]
        [DataRow("<table><tr><td class='custom'>First</td><td></td><td></td></tr></table>")]
        [DataRow("<table width='auto'><tr><td style='width:auto'>First</td><td></td><td></td></tr></table>")]
        [DataRow("<table><colgroup><col><col><col></colgroup><tr><td>First</td><td></td><td></td></tr></table>")]
        [DataRow("<table><tr><td colspan='2'>First</td><td></td></tr></table>")]
        [DataRow("<table><tr><td rowspan='2'>First</td><td></td><td></td></tr><tr><td>Second</td><td>Third</td></tr></table>")]
        [DataRow("<table><tr><td>First</td><td></td><td></td></tr><tr><td>Second</td><td></td><td></td></tr></table>")]
        [DataRow("<table><tr><th>Heading</th><th></th><th></th></tr><tr><td>First</td><td></td><td></td></tr></table>")]
        [DataRow("<table class='ms-rteTable-default'><tr><td class='ms-rteTableEvenCol-default'>First</td><td class='ms-rteTableOddCol-default'>Second</td><td>Third</td></tr></table>")]
        public void Transform_UnsizedTablesUseAutomaticColumnWidths(string html)
        {
            using (var document = new HtmlParser().ParseDocument(new HtmlTransformator().Transform(html, false)))
            {
                var cells = document.QuerySelectorAll("table tr > td");
                Assert.IsTrue(cells.Length >= 3);
                Assert.IsTrue(cells.All(cell => cell.GetAttribute("style").TrimEnd(';') == "width: auto"));
            }
        }

        [DataTestMethod]
        [DataRow("<table width='600'><tr><td>First</td><td></td><td></td></tr></table>")]
        [DataRow("<table style='width:90%'><tr><td>First</td><td></td><td></td></tr></table>")]
        [DataRow("<table><colgroup><col style='width:50%'></colgroup><tr><td>First</td><td></td><td></td></tr></table>")]
        [DataRow("<table><tr><td style='width:50%'>First</td><td></td><td></td></tr></table>")]
        [DataRow("<table><tr><td style='min-width:50px'>First</td><td></td><td></td></tr></table>")]
        [DataRow("<table><tr><td>First</td><td width='30'></td><td></td></tr></table>")]
        public void Transform_OtherTablesRetainDefaultColumnWidths(string html)
        {
            using (var document = new HtmlParser().ParseDocument(new HtmlTransformator().Transform(html, false)))
            {
                var result = document.QuerySelectorAll("table tr:first-child > td");
                CollectionAssert.AreEqual(new[] { "width: 266px", "width: 266px", "width: 268px" },
                    result.Select(cell => cell.GetAttribute("style").TrimEnd(';')).ToArray());
            }
        }

        [DataTestMethod]
        [DataRow(1, 2)]
        [DataRow(2, 2)]
        [DataRow(1, int.MaxValue)]
        [DataRow(2, int.MaxValue)]
        public void Transform_PreservesNestedTableContentWhenRowSpanExceedsRows(int colSpan, int rowSpan)
        {
            var html = "<div><table><tr><td colspan='3'><div>" +
                "<table class='ms-rteTable-7'><tr><td>" +
                "<h1>Welcome to the onboarding page</h1>" +
                "<h2>Important Security Updates</h2>" +
                "<p>CORP tenant registrations are unsupported.</p>" +
                "<p>Use a supported tenant.</p>" +
                "<ul><li><a href='/access-review'>Keycode access review</a></li></ul>" +
                $"</td><td colspan='{colSpan}' rowspan='{rowSpan}'>" +
                "<table><tr><td><p>Later security guidance</p></td></tr></table>" +
                "</td></tr></table></div></td></tr></table></div>";
            var observer = new Mock<ILogObserver>();
            var result = new HtmlTransformator(new[] { observer.Object }).Transform(html, false);

            using (var document = new HtmlParser().ParseDocument(result))
            {
                var expected = new[]
                {
                    "Welcome to the onboarding page",
                    "Important Security Updates",
                    "CORP tenant registrations are unsupported.",
                    "Use a supported tenant.",
                    "Keycode access review",
                    "Later security guidance"
                };
                CollectionAssert.AreEqual(expected,
                    document.QuerySelectorAll("h2, h3, p, li").Select(element => element.TextContent).ToArray());
                Assert.AreEqual("/access-review", document.QuerySelector("li > a").GetAttribute("href"));
                Assert.AreEqual(0, document.QuerySelectorAll("table table, [rowspan], [colspan]").Length);
                var introTable = document.QuerySelector("h2").Closest("table");
                Assert.AreEqual(1 + colSpan, introTable.QuerySelectorAll("td").Length);
            }
            observer.Verify(log => log.Warning(It.IsAny<LogEntry>()), Times.Never);
        }

        [DataTestMethod]
        [DataRow(1)]
        [DataRow(2)]
        public void Transform_ClipsRowSpanToRemainingRows(int colSpan)
        {
            var html = "<table><tr><td>First row</td></tr>" +
                $"<tr><td colspan='{colSpan}' rowspan='3'>Last row</td></tr></table>";
            var observer = new Mock<ILogObserver>();
            var result = new HtmlTransformator(new[] { observer.Object }).Transform(html, false);

            using (var document = new HtmlParser().ParseDocument(result))
            {
                Assert.AreEqual(1, document.QuerySelectorAll(".canvasRteResponsiveTable").Length);
                Assert.AreEqual(2, document.QuerySelectorAll("tr").Length);
                Assert.AreEqual(2 * colSpan, document.QuerySelectorAll("td").Length);
                Assert.AreEqual("First rowLast row", document.Body.TextContent);
                Assert.AreEqual(0, document.QuerySelectorAll("[rowspan], [colspan]").Length);
            }
            observer.Verify(log => log.Warning(It.IsAny<LogEntry>()), Times.Never);
        }

        [DataTestMethod]
        [DataRow(1)]
        [DataRow(2)]
        public void Transform_PreservesValidRowAndColumnSpans(int colSpan)
        {
            var html = $"<table><tr><th colspan='{colSpan + 1}'>Header</th></tr>" +
                $"<tr><td colspan='{colSpan}' rowspan='2'>Spanning cell</td><td>First</td></tr>" +
                "<tr><td>Second</td></tr></table>";
            var observer = new Mock<ILogObserver>();
            var result = new HtmlTransformator(new[] { observer.Object }).Transform(html, false);

            using (var document = new HtmlParser().ParseDocument(result))
            {
                var rows = document.QuerySelectorAll("tr");
                Assert.AreEqual(3, rows.Length);
                Assert.AreEqual("Header", rows[0].QuerySelector("td > strong").TextContent);
                Assert.AreEqual("Spanning cell", rows[1].Children[0].TextContent);
                Assert.AreEqual("First", rows[1].Children[colSpan].TextContent);
                Assert.IsTrue(rows[2].Children.Take(colSpan).All(cell => cell.TextContent == ""));
                Assert.AreEqual("Second", rows[2].Children[colSpan].TextContent);
                Assert.AreEqual(0, document.QuerySelectorAll("[rowspan], [colspan]").Length);
            }
            observer.Verify(log => log.Warning(It.IsAny<LogEntry>()), Times.Never);
        }
    }
}
