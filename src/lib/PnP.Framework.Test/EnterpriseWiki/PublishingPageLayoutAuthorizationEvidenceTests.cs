using Microsoft.VisualStudio.TestTools.UnitTesting;
using PnP.Framework.Migration.Evidence;
using PnP.Framework.Migration.Pages.Publishing.Assessment;
using PnP.Framework.Migration.Pages.Publishing.Capture;
using PnP.Framework.Migration.Pages.Publishing.Layouts;
using System;
using System.Linq;

namespace PnP.Framework.Test.EnterpriseWiki
{
    [TestClass]
    public class PublishingPageLayoutAuthorizationEvidenceTests
    {
        [TestMethod]
        public void ExternalLayoutHttp403BlocksOnlyItsCapturedIngredientBranch()
        {
            var wire = LiteralHttpAuthorizationEvidence.Create(
                "capture-page-layout-owner",
                "https://source.sharepoint.com/teams/staging/_vti_bin/client.svc/ProcessQuery",
                403,
                DateTimeOffset.Parse("2026-09-03T13:00:00Z"));
            var snapshot = new PublishingPageCaptureBundle
            {
                Layout = new PublishingPageLayoutSnapshot
                {
                    EvidenceState = PublishingPageLayoutEvidenceState.AuthorizationBlocked,
                    AuthorizationEvidence = wire
                }
            };

            var evidence = PublishingPageSnapshotAuthorizationEvidence.Merge(snapshot, null);

            Assert.AreEqual(2, evidence.AuthorizationFailures.Count);
            CollectionAssert.AreEquivalent(
                new[] { "layout:page", "content-type:page" },
                evidence.AuthorizationFailures.Select(value => value.IngredientId).ToArray());
            Assert.IsTrue(evidence.AuthorizationFailures.All(value => value.HttpStatusCode == 403));
            Assert.IsTrue(evidence.AuthorizationFailures.All(value => value.EvidenceSha256 == wire.EvidenceSha256));
        }
    }
}
