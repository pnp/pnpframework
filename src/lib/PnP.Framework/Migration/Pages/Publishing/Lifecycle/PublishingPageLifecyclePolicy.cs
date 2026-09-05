using System;

using PnP.Framework.Migration.Pages.Lifecycle;

namespace PnP.Framework.Migration.Pages.Publishing.Lifecycle
{
    internal static class PublishingPageLifecyclePolicy
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
