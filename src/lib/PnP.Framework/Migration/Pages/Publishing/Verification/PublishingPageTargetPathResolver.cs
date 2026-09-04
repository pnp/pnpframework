using PnP.Framework.Migration.Topology;
using System;
using System.Collections.Generic;

namespace PnP.Framework.Migration.Pages.Publishing.Verification
{
    internal sealed class PublishingPageTargetPathResolution
    {
        public string PreferredTargetPageServerRelativeUrl { get; set; }

        public string TargetPageServerRelativeUrl { get; set; }

        public bool CollisionResolved { get; set; }

        public string Reason { get; set; }
    }

    internal static class PublishingPageTargetPathResolver
    {
        public static PublishingPageTargetPathResolution Resolve(
            string preferredTargetPageServerRelativeUrl,
            string originalIdentifier,
            IEnumerable<string> occupiedPagePaths)
        {
            var target = TopologyTargetPathAllocator.AllocateServerRelativePath(
                preferredTargetPageServerRelativeUrl,
                originalIdentifier,
                occupiedPagePaths,
                preserveFileExtension: true);
            var changed = !string.Equals(target, preferredTargetPageServerRelativeUrl, StringComparison.Ordinal);
            return new PublishingPageTargetPathResolution
            {
                PreferredTargetPageServerRelativeUrl = preferredTargetPageServerRelativeUrl,
                TargetPageServerRelativeUrl = target,
                CollisionResolved = changed,
                Reason = changed
                    ? "Allocated a stable suffix only at the Page filename because the preferred target path is occupied."
                    : null
            };
        }
    }
}
