using AngleSharp.Html.Parser;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using PnP.Framework.Modernization.Transform;
using System.Linq;

namespace PnP.Framework.Modernization.Tests.Transform.Tables
{
    [TestClass]
    public class TableStyleTests
    {
        [DataTestMethod]
        [DataRow(2)]
        [DataRow(4)]
        [DataRow(5)]
        [DataRow(7)]
        [DataRow(8)]
        [DataRow(9)]
        [DataRow(10)]
        public void Transform_HeaderlessTableDoesNotFillFirstContentRow(int styleCode)
        {
            var html = $@"<table class='ms-rteTable-{styleCode}' cellspacing='0'>
                <tr class='ms-rteTableEvenRow-{styleCode}'>
                    <td class='ms-rteTableEvenCol-{styleCode}'>
                        <h2><span class='ms-rteForeColor-2'>Security requirements</span></h2>
                        <p>Read the <a href='/details'>instructions</a>.</p>
                    </td>
                </tr>
            </table>";

            using (var document = new HtmlParser().ParseDocument(new HtmlTransformator().Transform(html, false)))
            {
                var table = document.QuerySelector("table");
                Assert.AreEqual("simpleTableStyleNeutral", table.ClassName);
                Assert.AreEqual(1, table.QuerySelectorAll("tr").Length);
                Assert.AreEqual("Security requirements", table.QuerySelector("h3 .fontColorRed").TextContent);
                Assert.AreEqual("/details", table.QuerySelector("a").GetAttribute("href"));
                Assert.AreEqual("Read the instructions.", table.QuerySelector("p").TextContent);
            }
        }

        [DataTestMethod]
        [DataRow(2, "filledHeaderTableStyleNeutral")]
        [DataRow(4, "filledHeaderTableStyleTheme")]
        [DataRow(5, "filledHeaderTableStyleTheme")]
        [DataRow(7, "filledHeaderTableStyleTheme")]
        [DataRow(8, "filledHeaderTableStyleTheme")]
        [DataRow(9, "filledHeaderTableStyleTheme")]
        [DataRow(10, "filledHeaderTableStyleTheme")]
        public void Transform_PreservesFilledStyleForActualHeader(int styleCode, string expectedStyle)
        {
            var html = $@"<table class='ms-rteTable-{styleCode}'>
                <thead><tr><th>Heading</th><th>Details</th></tr></thead>
                <tbody><tr><td>First</td><td>Second</td></tr></tbody>
            </table>";

            using (var document = new HtmlParser().ParseDocument(new HtmlTransformator().Transform(html, false)))
            {
                var table = document.QuerySelector("table");
                Assert.AreEqual(expectedStyle, table.ClassName);
                Assert.AreEqual(2, table.QuerySelectorAll("tr").Length);
                CollectionAssert.AreEqual(new[] { "Heading", "Details" },
                    table.QuerySelectorAll("tr:first-child strong").Select(element => element.TextContent).ToArray());
                CollectionAssert.AreEqual(new[] { "First", "Second" },
                    table.QuerySelectorAll("tr:last-child td").Select(element => element.TextContent).ToArray());
            }
        }

        [DataTestMethod]
        [DataRow("<tr><td>First</td><td>Second</td></tr><tr><td>Third</td><td>Fourth</td></tr>")]
        [DataRow("<tr><th>First</th><td>Second</td></tr><tr><th>Third</th><td>Fourth</td></tr>")]
        [DataRow("<tr><td rowspan='2'>First</td><td>Second</td></tr><tr><td>Third</td></tr>")]
        [DataRow("<tr><td colspan='2'>First</td></tr><tr><td>Second</td><td>Third</td></tr>")]
        public void Transform_HeaderlessMultirowTableKeepsNeutralGrid(string rows)
        {
            var html = $"<table class='custom ms-rteTable-9' style='background-color:white;border:1px solid red'>{rows}</table>";

            using (var document = new HtmlParser().ParseDocument(new HtmlTransformator().Transform(html, false)))
            {
                var table = document.QuerySelector("table");
                Assert.AreEqual("simpleTableStyleNeutral", table.ClassName);
                Assert.AreEqual(2, table.QuerySelectorAll("tr").Length);
                Assert.AreEqual(4, table.QuerySelectorAll("td").Length);
                StringAssert.Contains(table.TextContent, "First");
                StringAssert.Contains(table.TextContent, "Second");
                StringAssert.Contains(table.TextContent, "Third");
            }
        }

        [TestMethod]
        public void Transform_NestedHeaderDoesNotGiveOuterTableAFilledHeader()
        {
            var html = @"<table class='ms-rteTable-7'><tr><td>Outer content
                <table class='ms-rteTable-9'>
                    <tr><th>Nested heading</th></tr><tr><td>Nested content</td></tr>
                </table>
            </td></tr></table>";

            using (var document = new HtmlParser().ParseDocument(new HtmlTransformator().Transform(html, false)))
            {
                var tables = document.QuerySelectorAll("table");
                Assert.AreEqual(2, tables.Length);
                Assert.AreEqual("simpleTableStyleNeutral", tables[0].ClassName);
                Assert.AreEqual("filledHeaderTableStyleTheme", tables[1].ClassName);
                StringAssert.Contains(tables[0].TextContent, "Outer content");
                Assert.AreEqual("Nested heading", tables[1].QuerySelector("strong").TextContent);
            }
        }

        [DataTestMethod]
        [DataRow("", "noBorderTableStyleNeutral")]
        [DataRow("class='ms-rteTable-0'", "noBorderTableStyleNeutral")]
        [DataRow("class='ms-rteTable-1'", "bandedRowTableStyleNeutral")]
        [DataRow("class='ms-rteTable-3'", "simpleTableStyleNeutral")]
        [DataRow("class='ms-rteTable-6'", "bandedRowTableStyleTheme")]
        [DataRow("class='ms-rteTable-default'", "borderHeaderTableStyleNeutral")]
        [DataRow("class='custom'", "borderHeaderTableStyleNeutral")]
        [DataRow("class='ms-rteTable-0' border='1'", "borderHeaderTableStyleNeutral")]
        [DataRow("style='border:1px solid red'", "borderHeaderTableStyleNeutral")]
        [DataRow("class='custom' border='0'", "noBorderTableStyleNeutral")]
        public void Transform_PreservesOtherTableStyles(string attributes, string expectedStyle)
        {
            var html = $"<table {attributes}><tr><td>Content</td></tr></table>";

            using (var document = new HtmlParser().ParseDocument(new HtmlTransformator().Transform(html, false)))
            {
                Assert.AreEqual(expectedStyle, document.QuerySelector("table").ClassName);
                Assert.AreEqual("Content", document.QuerySelector("td").TextContent);
            }
        }
    }
}
