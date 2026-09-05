using Microsoft.VisualStudio.TestTools.UnitTesting;
using PnP.Framework.Migration.Pages.Content;
using PnP.Framework.Migration.Pages.Publishing.Verification;
using System.Collections.Generic;

namespace PnP.Framework.Test.EnterpriseWiki
{
    [TestClass]
    public class PublishingPageStorageAndResumeIdentityTests
    {
        [TestMethod]
        public void EncodedUriSchemeColonIsEquivalentStorage()
        {
            var expected = "<img src=\"https://source.example/image.gif\" />";
            var persisted = "<img src=\"https&#58;//source.example/image.gif\" />";

            Assert.IsTrue(PublishingPageContentStorageCanonicalizer.AreEquivalent(expected, persisted));
            Assert.AreEqual(
                PublishingPageContentStorageCanonicalizer.ComputeCanonicalSha256(expected),
                PublishingPageContentStorageCanonicalizer.ComputeCanonicalSha256(persisted));
        }

        [TestMethod]
        public void CanonicalizerDoesNotDecodeEscapedMarkup()
        {
            Assert.IsFalse(PublishingPageContentStorageCanonicalizer.AreEquivalent(
                "<p>&lt;script&gt;</p>",
                "<p><script></p>"));
        }

        [TestMethod]
        public void PagesLibraryContentTypeDescendantMatchesApprovedSiteContentType()
        {
            const string expected = "0x010100AABB";
            const string actual = "0x010100AABB00112233445566778899AABBCCDDEEFF00";

            Assert.IsTrue(PublishingPageContentTypeIdentity.MatchesSiteContentType(actual, expected));
            Assert.IsTrue(PublishingPageContentTypeIdentity.MatchesSiteContentType(expected, expected));
            Assert.IsFalse(PublishingPageContentTypeIdentity.MatchesSiteContentType("0x010100AABC", expected));
        }

        [TestMethod]
        public void ResumeOwnershipRequiresAllSealedIdentities()
        {
            var properties = new Dictionary<string, object>
            {
                [PublishingPageTargetOwnership.OriginalIdentifierPropertyName] = "urn:page",
                [PublishingPageTargetOwnership.SourceSnapshotDigestPropertyName] = "source-digest",
                [PublishingPageTargetOwnership.PlanDigestPropertyName] = "plan-digest"
            };

            Assert.IsTrue(PublishingPageTargetOwnership.MatchesApprovedPlan(
                properties,
                "urn:page",
                "SOURCE-DIGEST",
                "PLAN-DIGEST"));
            Assert.IsFalse(PublishingPageTargetOwnership.MatchesApprovedPlan(
                properties,
                "urn:other",
                "source-digest",
                "plan-digest"));
        }
    }
}
