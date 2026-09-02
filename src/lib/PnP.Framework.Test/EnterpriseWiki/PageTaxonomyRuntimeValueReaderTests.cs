using Microsoft.VisualStudio.TestTools.UnitTesting;
using PnP.Framework.Migration.Pages.Fields.Taxonomy;
using System.Collections.Generic;
using System.Linq;

namespace PnP.Framework.Test.EnterpriseWiki
{
    [TestClass]
    public class PageTaxonomyRuntimeValueReaderTests
    {
        [TestMethod]
        public void ReadsBearerAuthenticatedCsomTaxonomyCollectionDictionary()
        {
            var runtimeValue = new Dictionary<string, object>
            {
                ["_ObjectType_"] = "SP.Taxonomy.TaxonomyFieldValueCollection",
                ["_Child_Items_"] = new object[]
                {
                    new Dictionary<string, object>
                    {
                        ["_ObjectType_"] = "SP.Taxonomy.TaxonomyFieldValue",
                        ["Label"] = "Resource Library",
                        ["TermGuid"] = "30a6305a-4903-41c7-b2d9-9793f1c1cf4a",
                        ["WssId"] = 35
                    }
                }
            };

            Assert.IsTrue(PageTaxonomyRuntimeValueReader.TryRead(runtimeValue, out var result));
            Assert.IsTrue(result.IsCollection);
            Assert.AreEqual(1, result.Values.Count);
            Assert.AreEqual("Resource Library", result.Values.Single().Label);
            Assert.AreEqual("30a6305a-4903-41c7-b2d9-9793f1c1cf4a", result.Values.Single().TermGuid);
            Assert.AreEqual(35, result.Values.Single().WssId);
        }

        [TestMethod]
        public void PreservesMalformedTaxonomyIdentityForConflictClassification()
        {
            var runtimeValue = new Dictionary<string, object>
            {
                ["_ObjectType_"] = "SP.Taxonomy.TaxonomyFieldValue",
                ["Label"] = "Broken",
                ["TermGuid"] = "not-a-guid",
                ["WssId"] = "17"
            };

            Assert.IsTrue(PageTaxonomyRuntimeValueReader.TryRead(runtimeValue, out var result));
            Assert.IsFalse(result.IsCollection);
            Assert.AreEqual("not-a-guid", result.Values.Single().TermGuid);
            Assert.AreEqual(17, result.Values.Single().WssId);
        }
    }
}
