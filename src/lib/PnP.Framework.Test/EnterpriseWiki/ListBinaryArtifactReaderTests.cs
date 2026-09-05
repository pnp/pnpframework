using Microsoft.VisualStudio.TestTools.UnitTesting;
using PnP.Framework.Migration.Lists.Capture;
using PnP.Framework.Migration.Lists.Items;
using PnP.Framework.Migration.Lists.Items.Protection;
using PnP.Framework.Migration.Pages.Publishing.Packaging;
using System.Collections.Generic;
using System.Text;

namespace PnP.Framework.Test.EnterpriseWiki
{
    [TestClass]
    public class ListBinaryArtifactReaderTests
    {
        [DataTestMethod]
        [DataRow("fullyArchived")]
        [DataRow("FullyArchived")]
        [DataRow(" fullyArchived ")]
        public void FullyArchivedMetadataSelectsArchiveFastPath(string value)
        {
            Assert.IsTrue(ListBinaryArtifactReader.IsFullyArchived(value));
        }

        [DataTestMethod]
        [DataRow(null)]
        [DataRow("")]
        [DataRow("archived")]
        [DataRow("reactivating")]
        [DataRow("active")]
        public void OtherMetadataDoesNotSelectArchiveFastPath(string value)
        {
            Assert.IsFalse(ListBinaryArtifactReader.IsFullyArchived(value));
        }

        [TestMethod]
        public void DrmCompoundPayloadIsClassifiedAsRightsManagedEnvelope()
        {
            var payload = new byte[1024];
            new byte[] { 0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1 }
                .CopyTo(payload, 0);
            Encoding.Unicode.GetBytes("DRMEncryptedTransform").CopyTo(payload, 128);
            Encoding.Unicode.GetBytes("EncryptedPackage").CopyTo(payload, 512);

            Assert.AreEqual(
                ListBinaryRepresentationKind.InformationRightsManagedEnvelope,
                ListBinaryPayloadClassifier.Classify(payload));
        }

        [TestMethod]
        public void CompoundPayloadWithoutBothDrmMarkersRemainsOrdinary()
        {
            var payload = new byte[512];
            new byte[] { 0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1 }
                .CopyTo(payload, 0);
            Encoding.Unicode.GetBytes("EncryptedPackage").CopyTo(payload, 128);

            Assert.AreEqual(
                ListBinaryRepresentationKind.OrdinaryFilePayload,
                ListBinaryPayloadClassifier.Classify(payload));
        }

        [TestMethod]
        public void LegacyUnclassifiedBinaryOmitsNewRepresentationFieldFromCanonicalPayload()
        {
            var legacy = PublishingPagePackageSerializer.SerializeCanonical(
                new ListBinaryArtifactSnapshot());
            var current = PublishingPagePackageSerializer.SerializeCanonical(
                new ListBinaryArtifactSnapshot
                {
                    RepresentationKind = ListBinaryRepresentationKind.OrdinaryFilePayload
                });

            Assert.IsFalse(legacy.Contains("representationKind", System.StringComparison.Ordinal));
            StringAssert.Contains(current, "\"representationKind\":\"OrdinaryFilePayload\"");
        }

        [TestMethod]
        public void LegacyListSnapshotOmitsNewInformationRightsManagementContract()
        {
            var legacy = PublishingPagePackageSerializer.SerializeCanonical(
                new ListDependencySnapshot());
            var current = PublishingPagePackageSerializer.SerializeCanonical(
                new ListDependencySnapshot
                {
                    InformationRightsManagement = new ListInformationRightsManagementSnapshot
                    {
                        IrmEnabled = true,
                        Policy = new ListInformationRightsManagementPolicySnapshot
                        {
                            PolicyTitle = "Protected library"
                        }
                    }
                });

            Assert.IsFalse(legacy.Contains("informationRightsManagement", System.StringComparison.Ordinal));
            StringAssert.Contains(current, "\"irmEnabled\":true");
            StringAssert.Contains(current, "\"policyTitle\":\"Protected library\"");
        }

        [TestMethod]
        public void InformationProtectionReaderProjectsExactExternalLabelEvidence()
        {
            var values = new Dictionary<string, object>
            {
                ["_IpLabelId"] = "9fbde396-1a24-4c79-8edf-9254a0f35055",
                ["_IpLabelAssignmentMethod"] = "1",
                ["_HasUserDefinedProtection"] = "0",
                ["_IpLabelOwnerEmail"] = "owner@source.example",
                ["_IpLabelHash"] = "label-hash;00",
                ["_IpLabelPromotionCtagVersion"] = "20",
                ["MetaInfo"] = "vti_title:SW|Protected deck\r\nvti_decryptskipreason:IW|1\r\n"
            };

            var snapshot = ListDocumentInformationProtectionSnapshotReader.Read(values);

            Assert.IsNotNull(snapshot);
            Assert.AreEqual("9fbde396-1a24-4c79-8edf-9254a0f35055", snapshot.LabelId);
            Assert.AreEqual("1", snapshot.AssignmentMethod);
            Assert.AreEqual("0", snapshot.HasUserDefinedProtection);
            Assert.AreEqual("owner@source.example", snapshot.OwnerEmail);
            Assert.AreEqual("label-hash;00", snapshot.LabelHash);
            Assert.AreEqual("20", snapshot.PromotionCtagVersion);
            Assert.AreEqual("1", snapshot.DecryptSkipReason);
        }

        [TestMethod]
        public void InformationProtectionReaderDoesNotInventMissingAssignment()
        {
            Assert.IsNull(ListDocumentInformationProtectionSnapshotReader.Read(
                new Dictionary<string, object>
                {
                    ["_IpLabelId"] = string.Empty,
                    ["_HasUserDefinedProtection"] = "0"
                }));
        }

        [TestMethod]
        public void MediaServiceMetadataProvidesStableLogicalContentIdentity()
        {
            var values = new Dictionary<string, object>
            {
                ["metainfo"] = "unrelated:SW|value\r\n"
                    + "vti_mediaservicemetadata:SW|{\"ctag\":\"c:123\",\"quickxorhash\":\"YWJjZA==\"}\r\n"
            };

            var identity = ListBinaryContentIdentityReader.Read(values);

            Assert.IsNotNull(identity);
            Assert.AreEqual("YWJjZA==", identity.QuickXorHash);
            Assert.AreEqual("c:123", identity.ContentTag);
            Assert.AreEqual("SharePoint.MetaInfo.vti_mediaservicemetadata", identity.EvidenceSource);
        }

        [TestMethod]
        public void DirectMediaServiceMetadataProvidesStableLogicalContentIdentity()
        {
            var values = new Dictionary<string, object>
            {
                ["MediaServiceMetadata"] = "{\"ctag\":\"\\\"c:{671F67F4-D605-41FB-B93D-CFF280AD13C4},20\\\"\",\"quickxorhash\":\"wntwbzbUjfGnS9YwLJjIvJZyndY=\"}",
                ["MetaInfo"] = "vti_mediaservicemetadata:SW|{not-json}"
            };

            var identity = ListBinaryContentIdentityReader.Read(values);

            Assert.IsNotNull(identity);
            Assert.AreEqual("wntwbzbUjfGnS9YwLJjIvJZyndY=", identity.QuickXorHash);
            Assert.AreEqual("\"c:{671F67F4-D605-41FB-B93D-CFF280AD13C4},20\"", identity.ContentTag);
            Assert.AreEqual("SharePoint.ListItem.MediaServiceMetadata", identity.EvidenceSource);
        }

        [TestMethod]
        public void MalformedDirectMediaServiceMetadataFallsBackToMetaInfo()
        {
            var values = new Dictionary<string, object>
            {
                ["MediaServiceMetadata"] = "{not-json}",
                ["MetaInfo"] = "vti_mediaservicemetadata:SW|{\"ctag\":\"c:123\",\"quickxorhash\":\"YWJjZA==\"}"
            };

            var identity = ListBinaryContentIdentityReader.Read(values);

            Assert.IsNotNull(identity);
            Assert.AreEqual("YWJjZA==", identity.QuickXorHash);
            Assert.AreEqual("SharePoint.MetaInfo.vti_mediaservicemetadata", identity.EvidenceSource);
        }

        [TestMethod]
        public void MalformedMediaServiceMetadataDoesNotInventIdentity()
        {
            var values = new Dictionary<string, object>
            {
                ["MetaInfo"] = "vti_mediaservicemetadata:SW|{not-json}"
            };

            Assert.IsNull(ListBinaryContentIdentityReader.Read(values));
        }
    }
}
