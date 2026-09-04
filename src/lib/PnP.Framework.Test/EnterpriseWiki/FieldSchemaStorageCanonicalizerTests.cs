using Microsoft.VisualStudio.TestTools.UnitTesting;
using PnP.Framework.Migration.Schema.Fields;

namespace PnP.Framework.Test.EnterpriseWiki
{
    [TestClass]
    public class FieldSchemaStorageCanonicalizerTests
    {
        [TestMethod]
        public void PortableSchemaTreatsExpandedServerDefaultsAsEquivalentToOmission()
        {
            var compact = "<Field Type=\"Text\" DisplayName=\"Derived from ID\" Description=\"Source identity\" Required=\"FALSE\" "
                + "EnforceUniqueValues=\"FALSE\" Indexed=\"FALSE\" MaxLength=\"255\" Group=\"Shared\" "
                + "ID=\"{7e229256-a9ae-4bbe-a70f-391a14d95dcc}\" StaticName=\"DerivedFromID\" Name=\"DerivedFromID\">"
                + "<Default>Original</Default></Field>";
            var expanded = "<Field Type=\"Text\" DisplayName=\"Derived from ID\" Description=\"Source identity\" Required=\"FALSE\" "
                + "EnforceUniqueValues=\"FALSE\" Indexed=\"FALSE\" MaxLength=\"255\" Group=\"Shared\" "
                + "ID=\"{7e229256-a9ae-4bbe-a70f-391a14d95dcc}\" StaticName=\"DerivedFromID\" Name=\"DerivedFromID\" "
                + "Hidden=\"FALSE\" ReadOnly=\"FALSE\" Sealed=\"FALSE\" Customization=\"\" PITarget=\"\" PrimaryPITarget=\"\" "
                + "PIAttribute=\"\" PrimaryPIAttribute=\"\" Aggregation=\"\" Node=\"\" AllowDeletion=\"TRUE\">"
                + "<Default>Original</Default></Field>";

            Assert.AreEqual(
                FieldSchemaCanonicalizer.PortableDigest(compact),
                FieldSchemaCanonicalizer.PortableDigest(expanded));
        }

        [TestMethod]
        public void PortableSchemaRetainsMaterialVisibilityAndDefaultValueDifferences()
        {
            const string baseline = "<Field Type=\"Text\" ID=\"{7e229256-a9ae-4bbe-a70f-391a14d95dcc}\" Name=\"DerivedFromID\"><Default>Original</Default></Field>";
            const string hidden = "<Field Type=\"Text\" ID=\"{7e229256-a9ae-4bbe-a70f-391a14d95dcc}\" Name=\"DerivedFromID\" Hidden=\"TRUE\"><Default>Original</Default></Field>";
            const string changedDefault = "<Field Type=\"Text\" ID=\"{7e229256-a9ae-4bbe-a70f-391a14d95dcc}\" Name=\"DerivedFromID\"><Default>Changed</Default></Field>";

            Assert.AreNotEqual(FieldSchemaCanonicalizer.PortableDigest(baseline), FieldSchemaCanonicalizer.PortableDigest(hidden));
            Assert.AreNotEqual(FieldSchemaCanonicalizer.PortableDigest(baseline), FieldSchemaCanonicalizer.PortableDigest(changedDefault));
        }
    }
}
