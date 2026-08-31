using System;

namespace PnP.Framework.Migration.PublishingPages
{
    internal static class PageLifecyclePolicy
    {
        public static PublishingPageTargetLifecycle DeriveTargetLifecycle(PageLifecycleSnapshot sourceLifecycle)
        {
            var isPublished = string.Equals(sourceLifecycle?.Level, "Published", StringComparison.OrdinalIgnoreCase)
                && string.Equals(sourceLifecycle?.CheckOutType, "None", StringComparison.OrdinalIgnoreCase)
                && (!sourceLifecycle.ModerationStatus.HasValue || sourceLifecycle.ModerationStatus.Value == 0);
            return isPublished
                ? PublishingPageTargetLifecycle.Published
                : PublishingPageTargetLifecycle.Draft;
        }
    }
}
