using PnP.Framework.Migration.PublishingPages.Capture;
using PnP.Framework.Migration.PublishingPages.Lifecycle;
using PnP.Framework.Migration.PublishingPages.References;
using System.Collections.Generic;
using System.Linq;

namespace PnP.Framework.Migration.PublishingPages.Verification
{
    internal static class PageStorageAssertionBuilder
    {
        public static IList<string> Build(
            PublishingPageCaptureBundle snapshot,
            string targetPagePath,
            IEnumerable<PageReferenceAction> referenceActions,
            string expectedContentDigest,
            PublishingPageTargetLifecycle targetLifecycle)
        {
            var result = new List<string>
            {
                $"target-page={targetPagePath}",
                "fresh-read-target-file-identity",
                "fresh-read-target-page-content-type",
                "fresh-read-target-version-and-lifecycle",
                $"expected-target-lifecycle={targetLifecycle}",
                $"source-publishing-page-content-sha256={snapshot.PublishingPageContentSha256}",
                $"expected-target-publishing-page-content-sha256={expectedContentDigest}",
                $"expected-shared-webparts={snapshot.WebParts.Count}"
            };
            var referenceById = snapshot.Dependencies.ToDictionary(item => item.Id, System.StringComparer.Ordinal);
            result.AddRange(referenceActions
                .Where(item => item.Disposition == PageReferenceDisposition.MaterializeAtTarget)
                .Select(item => $"dependency={item.TargetServerRelativeUrl}|sha256={referenceById[item.SnapshotDependencyId].ContentSha256}"));
            return result.OrderBy(item => item, System.StringComparer.Ordinal).ToList();
        }
    }
}
