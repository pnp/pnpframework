using PnP.Framework.Migration.PublishingPages.Fields;
using PnP.Framework.Migration.PublishingPages.Lifecycle;
using PnP.Framework.Migration.PublishingPages.References;
using PnP.Framework.Migration.PublishingPages.Security;
using PnP.Framework.Migration.PublishingPages.WebParts;
using System.Collections.Generic;

namespace PnP.Framework.Migration.PublishingPages.Capture
{
    public sealed class PublishingPageCaptureBundle
    {
        public string SourceProfile { get; set; }

        public PublishingPageExportOptions CapturePolicy { get; set; }

        public PublishingPageIdentity Source { get; set; }

        public string PublishingPageContent { get; set; }

        public string PublishingPageContentSha256 { get; set; }

        public IList<PageFieldValueSnapshot> Fields { get; set; } = new List<PageFieldValueSnapshot>();

        public IList<PageWebPartSnapshot> WebParts { get; set; } = new List<PageWebPartSnapshot>();

        public IList<PageReferenceSnapshot> Dependencies { get; set; } = new List<PageReferenceSnapshot>();

        public PageSecuritySnapshot Security { get; set; }

        public PageLifecycleSnapshot Lifecycle { get; set; }

        public SourcePageFence SourceFence { get; set; }

        public IList<string> Blockers { get; set; } = new List<string>();

        public IList<string> Warnings { get; set; } = new List<string>();
    }
}
