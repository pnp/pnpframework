using PnP.Framework.Migration.Pages.Capture;
using PnP.Framework.Migration.Pages.Fields;
using PnP.Framework.Migration.Pages.Lifecycle;
using PnP.Framework.Migration.Pages.References;
using PnP.Framework.Migration.Pages.Security;
using PnP.Framework.Migration.Pages.ClassicWebParts;
using System.Collections.Generic;

namespace PnP.Framework.Migration.Pages.Publishing.Capture
{
    public sealed class PublishingPageCaptureBundle
    {
        public string SourceProfile { get; set; }

        public PageCaptureOptions CapturePolicy { get; set; }

        public PageIdentity Source { get; set; }

        public PublishingPageLayoutSnapshot Layout { get; set; }

        public string PublishingPageContent { get; set; }

        public string PublishingPageContentSha256 { get; set; }

        public IList<PageFieldValueSnapshot> Fields { get; set; } = new List<PageFieldValueSnapshot>();

        public IList<ClassicWebPartSnapshot> WebParts { get; set; } = new List<ClassicWebPartSnapshot>();

        public IList<PageReferenceSnapshot> Dependencies { get; set; } = new List<PageReferenceSnapshot>();

        public PageSecuritySnapshot Security { get; set; }

        public PageLifecycleSnapshot Lifecycle { get; set; }

        public SourcePageFence SourceFence { get; set; }

        public IList<string> Blockers { get; set; } = new List<string>();

        public IList<string> Warnings { get; set; } = new List<string>();
    }
}
