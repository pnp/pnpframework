using PnP.Framework.Migration.Pages.Capture;
using PnP.Framework.Migration.Pages.Fields;
using PnP.Framework.Migration.Pages.Lifecycle;
using PnP.Framework.Migration.Pages.Security;
using PnP.Framework.Migration.Pages.ClassicWebParts;
using System.Collections.Generic;

namespace PnP.Framework.Migration.Pages.Publishing.Capture
{
    internal sealed class CapturedPublishingPage
    {
        public PageIdentity Identity { get; set; }

        public PublishingPageLayoutSnapshot Layout { get; set; }

        public string PublishingPageContent { get; set; }

        public List<PageFieldValueSnapshot> Fields { get; set; }

        public List<ClassicWebPartSnapshot> WebParts { get; set; }

        public PageSecuritySnapshot Security { get; set; }

        public PageLifecycleSnapshot Lifecycle { get; set; }

        public SourcePageFence SourceFence { get; set; }
    }
}
