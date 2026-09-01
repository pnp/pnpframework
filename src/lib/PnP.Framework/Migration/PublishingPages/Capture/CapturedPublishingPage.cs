using PnP.Framework.Migration.PublishingPages.Fields;
using PnP.Framework.Migration.PublishingPages.Lifecycle;
using PnP.Framework.Migration.PublishingPages.Security;
using PnP.Framework.Migration.PublishingPages.WebParts;
using System.Collections.Generic;

namespace PnP.Framework.Migration.PublishingPages.Capture
{
    internal sealed class CapturedPublishingPage
    {
        public PublishingPageIdentity Identity { get; set; }

        public string PublishingPageContent { get; set; }

        public List<PageFieldValueSnapshot> Fields { get; set; }

        public List<PageWebPartSnapshot> WebParts { get; set; }

        public PageSecuritySnapshot Security { get; set; }

        public PageLifecycleSnapshot Lifecycle { get; set; }

        public SourcePageFence SourceFence { get; set; }
    }
}
