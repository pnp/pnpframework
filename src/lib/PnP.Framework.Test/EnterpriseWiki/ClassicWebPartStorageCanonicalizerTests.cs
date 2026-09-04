using Microsoft.VisualStudio.TestTools.UnitTesting;
using PnP.Framework.Migration.Pages.ClassicWebParts.Bindings;

namespace PnP.Framework.Test.EnterpriseWiki
{
    [TestClass]
    public class ClassicWebPartStorageCanonicalizerTests
    {
        [TestMethod]
        public void ListBoundExportTreatsRuntimeViewIdentityAndEquivalentEmptyRepresentationAsStorageLocal()
        {
            var expected = Export(
                "11111111-1111-1111-1111-111111111111",
                "{22222222-2222-2222-2222-222222222222}",
                "<property name=\"NoDefaultStyle\" type=\"string\" />");
            var actual = Export(
                "11111111-1111-1111-1111-111111111111",
                "{33333333-3333-3333-3333-333333333333}",
                "<property name=\"NoDefaultStyle\" type=\"string\" null=\"true\" />")
                .Replace("{AAAAAAAA-AAAA-AAAA-AAAA-AAAAAAAAAAAA}", "{aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa}")
                .Replace("https://source.example/_layouts/15/images/links.png?rev=50", "/_layouts/15/images/links.png?rev=50");

            Assert.AreEqual(
                ClassicWebPartStorageCanonicalizer.CanonicalizeListBoundExport(expected),
                ClassicWebPartStorageCanonicalizer.CanonicalizeListBoundExport(actual));
        }

        [TestMethod]
        public void ListBoundExportRetainsMaterialListIdentityAndCamlDifferences()
        {
            var expected = Export(
                "11111111-1111-1111-1111-111111111111",
                "{22222222-2222-2222-2222-222222222222}",
                "<property name=\"NoDefaultStyle\" type=\"string\" />");
            var wrongList = Export(
                "99999999-9999-9999-9999-999999999999",
                "{33333333-3333-3333-3333-333333333333}",
                "<property name=\"NoDefaultStyle\" type=\"string\" null=\"true\" />");

            Assert.AreNotEqual(
                ClassicWebPartStorageCanonicalizer.CanonicalizeListBoundExport(expected),
                ClassicWebPartStorageCanonicalizer.CanonicalizeListBoundExport(wrongList));
        }

        [TestMethod]
        public void ListBoundExportNormalizesDuplicatedOriginBeforeBuiltInLayoutIcon()
        {
            var expected = Export(
                "11111111-1111-1111-1111-111111111111",
                "{22222222-2222-2222-2222-222222222222}",
                "<property name=\"NoDefaultStyle\" type=\"string\" />")
                .Replace(
                    "https://source.example/_layouts/15/images/links.png?rev=50",
                    "https://source.examplehttps://source.example/_layouts/15/images/links.png?rev=50");
            var actual = Export(
                "11111111-1111-1111-1111-111111111111",
                "{33333333-3333-3333-3333-333333333333}",
                "<property name=\"NoDefaultStyle\" type=\"string\" null=\"true\" />")
                .Replace("https://source.example/_layouts/15/images/links.png?rev=50", "/_layouts/15/images/links.png?rev=50");

            Assert.AreEqual(
                ClassicWebPartStorageCanonicalizer.CanonicalizeListBoundExport(expected),
                ClassicWebPartStorageCanonicalizer.CanonicalizeListBoundExport(actual));
        }

        private static string Export(string listId, string viewId, string emptyProperty) =>
            "<webParts><webPart xmlns=\"http://schemas.microsoft.com/WebPart/v3\"><metaData>"
            + "<type name=\"Microsoft.SharePoint.WebPartPages.XsltListViewWebPart\" /></metaData><data><properties>"
            + emptyProperty
            + "<property name=\"ListId\">" + listId + "</property>"
            + "<property name=\"ListName\">{AAAAAAAA-AAAA-AAAA-AAAA-AAAAAAAAAAAA}</property>"
            + "<property name=\"WebId\">bbbbbbbb-bbbb-bbbb-bbbb-bbbbbbbbbbbb</property>"
            + "<property name=\"XmlDefinition\">&lt;View Name=\"" + viewId
            + "\" Url=\"/sites/target/Pages/A.aspx\" ImageUrl=\"https://source.example/_layouts/15/images/links.png?rev=50\"&gt;"
            + "&lt;ViewFields&gt;&lt;FieldRef Name=\"Title\" /&gt;&lt;/ViewFields&gt;&lt;/View&gt;</property>"
            + "</properties></data></webPart></webParts>";
    }
}
