using Microsoft.VisualStudio.TestTools.UnitTesting;
using PnP.Framework.Migration.Schema.Fields;
using System;
using System.IO;
using System.Linq;

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
        public void PortableSchemaTreatsServerRequiredFalseAsEquivalentToOmission()
        {
            const string compact = "<Field Type=\"TaxonomyFieldType\" DisplayName=\"Language\" "
                + "ID=\"{92cc2700-0556-43d8-b753-8b146d9b1b61}\" Name=\"MS_x0020_Language\" />";
            const string expanded = "<Field Type=\"TaxonomyFieldType\" DisplayName=\"Language\" Required=\"FALSE\" "
                + "ID=\"{92cc2700-0556-43d8-b753-8b146d9b1b61}\" Name=\"MS_x0020_Language\" />";

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

        [TestMethod]
        public void PortableSchemaIgnoresRuntimeDisplayNameForHiddenTaxonomyCompanionOnly()
        {
            const string sourceCompanion = "<Field Type=\"Note\" DisplayName=\"ServicesDomain_0\" Hidden=\"TRUE\" "
                + "ID=\"{35fe1c89-1f26-4c87-bd68-061ced1afdb3}\" Name=\"g6775e77a6d84637a29014d883a4378a\" />";
            const string targetCompanion = "<Field Type=\"Note\" DisplayName=\"Services Domain_0\" Hidden=\"TRUE\" "
                + "ID=\"{35fe1c89-1f26-4c87-bd68-061ced1afdb3}\" Name=\"g6775e77a6d84637a29014d883a4378a\" />";
            const string ordinaryField = "<Field Type=\"Note\" DisplayName=\"Services Domain_0\" Hidden=\"FALSE\" "
                + "ID=\"{35fe1c89-1f26-4c87-bd68-061ced1afdb3}\" Name=\"VisibleNotes\" />";
            const string renamedOrdinaryField = "<Field Type=\"Note\" DisplayName=\"Different\" Hidden=\"FALSE\" "
                + "ID=\"{35fe1c89-1f26-4c87-bd68-061ced1afdb3}\" Name=\"VisibleNotes\" />";

            Assert.AreEqual(
                FieldSchemaCanonicalizer.PortableDigest(sourceCompanion),
                FieldSchemaCanonicalizer.PortableDigest(targetCompanion));
            Assert.AreNotEqual(
                FieldSchemaCanonicalizer.PortableDigest(ordinaryField),
                FieldSchemaCanonicalizer.PortableDigest(renamedOrdinaryField));
        }

        [TestMethod]
        public void TargetRewritePreservesHiddenTaxonomyCompanionDisplayName()
        {
            const string source = "<Field Type=\"Note\" DisplayName=\"ServicesDomain_0\" Hidden=\"TRUE\" "
                + "ShowInViewForms=\"FALSE\" CanToggleHidden=\"TRUE\" SourceID=\"{11111111-1111-1111-1111-111111111111}\" "
                + "ID=\"{35fe1c89-1f26-4c87-bd68-061ced1afdb3}\" Name=\"g6775e77a6d84637a29014d883a4378a\" />";

            var target = FieldSchemaCanonicalizer.RewriteForTarget(source);

            StringAssert.Contains(target, "DisplayName=\"ServicesDomain_0\"");
            Assert.AreEqual(
                FieldSchemaCanonicalizer.PortableDigest(source),
                FieldSchemaCanonicalizer.PortableDigest(target));
        }

        [TestMethod]
        public void PortableSchemaAcceptsRuntimeIndexOnDocumentIdValueFieldOnly()
        {
            const string source = "<Field Type=\"Text\" ID=\"{ae3e2a36-125d-45d3-9051-744b513536a6}\" Name=\"_dlc_DocId\" />";
            const string runtime = "<Field Type=\"Text\" ID=\"{ae3e2a36-125d-45d3-9051-744b513536a6}\" Name=\"_dlc_DocId\" Indexed=\"TRUE\" />";
            const string ordinarySource = "<Field Type=\"Text\" ID=\"{11111111-1111-1111-1111-111111111111}\" Name=\"Value\" />";
            const string ordinaryIndexed = "<Field Type=\"Text\" ID=\"{11111111-1111-1111-1111-111111111111}\" Name=\"Value\" Indexed=\"TRUE\" />";

            Assert.AreEqual(FieldSchemaCanonicalizer.PortableDigest(source), FieldSchemaCanonicalizer.PortableDigest(runtime));
            Assert.AreNotEqual(FieldSchemaCanonicalizer.PortableDigest(ordinarySource), FieldSchemaCanonicalizer.PortableDigest(ordinaryIndexed));
        }

        [TestMethod]
        public void SiteFieldMergeSelectsParentProducerForMatchingInheritedConsumer()
        {
            var producer = Plan(
                FieldSchemaRole.DirectBinding,
                FieldOwnership.UserDefined,
                FieldSchemaMaterializationDisposition.CreateOrReuseOwned,
                "<Field Type=\"Text\" ID=\"{7e229256-a9ae-4bbe-a70f-391a14d95dcc}\" Name=\"DerivedFromID\" />");
            var consumer = Plan(
                FieldSchemaRole.InheritedFromParent,
                FieldOwnership.TargetRuntime,
                FieldSchemaMaterializationDisposition.RequireTargetRuntime,
                null);

            Assert.AreSame(
                producer,
                SiteFieldMaterializer.Merge(producer.FieldId, new[] { consumer, producer }));
        }

        [TestMethod]
        public void SiteFieldMergeRejectsInheritedConsumerWithDifferentSourceSchema()
        {
            var producer = Plan(
                FieldSchemaRole.DirectBinding,
                FieldOwnership.UserDefined,
                FieldSchemaMaterializationDisposition.CreateOrReuseOwned,
                "<Field Type=\"Text\" ID=\"{7e229256-a9ae-4bbe-a70f-391a14d95dcc}\" Name=\"DerivedFromID\" />");
            var consumer = Plan(
                FieldSchemaRole.InheritedFromParent,
                FieldOwnership.TargetRuntime,
                FieldSchemaMaterializationDisposition.RequireTargetRuntime,
                null);
            consumer.SourcePortableSchemaSha256 = "different";

            Assert.ThrowsException<InvalidDataException>(() =>
                SiteFieldMaterializer.Merge(producer.FieldId, new[] { producer, consumer }));
        }

        [TestMethod]
        public void SiteFieldMaterializationOrdersHiddenCompanionBeforeTaxonomyConsumer()
        {
            var companion = Plan(
                FieldSchemaRole.DirectBinding,
                FieldOwnership.UserDefined,
                FieldSchemaMaterializationDisposition.CreateOrReuseOwned,
                "<Field Type=\"Note\" ID=\"{35fe1c89-1f26-4c87-bd68-061ced1afdb3}\" Name=\"g6775e77a6d84637a29014d883a4378a\" />");
            companion.FieldId = Guid.Parse("35fe1c89-1f26-4c87-bd68-061ced1afdb3");
            companion.InternalName = "g6775e77a6d84637a29014d883a4378a";
            companion.TypeAsString = "Note";
            var taxonomy = Plan(
                FieldSchemaRole.DirectBinding,
                FieldOwnership.UserDefined,
                FieldSchemaMaterializationDisposition.CreateOrReuseOwned,
                "<Field Type=\"TaxonomyFieldTypeMulti\" ID=\"{06775e77-a6d8-4637-a290-14d883a4378a}\" Name=\"ServicesDomain\" />");
            taxonomy.FieldId = Guid.Parse("06775e77-a6d8-4637-a290-14d883a4378a");
            taxonomy.InternalName = "ServicesDomain";
            taxonomy.TypeAsString = "TaxonomyFieldTypeMulti";
            taxonomy.HiddenTextFieldId = companion.FieldId;

            var ordered = SiteFieldMaterializer.OrderForMaterialization(new[] { taxonomy, companion }).ToArray();

            Assert.AreSame(companion, ordered[0]);
            Assert.AreSame(taxonomy, ordered[1]);
        }

        private static FieldSchemaMaterializationPlan Plan(
            FieldSchemaRole role,
            FieldOwnership ownership,
            FieldSchemaMaterializationDisposition disposition,
            string targetSchema)
        {
            const string source = "<Field Type=\"Text\" ID=\"{7e229256-a9ae-4bbe-a70f-391a14d95dcc}\" Name=\"DerivedFromID\" />";
            return new FieldSchemaMaterializationPlan
            {
                FieldId = Guid.Parse("7e229256-a9ae-4bbe-a70f-391a14d95dcc"),
                InternalName = "DerivedFromID",
                TypeAsString = "Text",
                Role = role,
                Ownership = ownership,
                Disposition = disposition,
                SourcePortableSchemaSha256 = FieldSchemaCanonicalizer.PortableDigest(source),
                TargetSchemaXml = targetSchema,
                TargetPortableSchemaSha256 = targetSchema == null
                    ? null
                    : FieldSchemaCanonicalizer.PortableDigest(targetSchema)
            };
        }
    }
}
