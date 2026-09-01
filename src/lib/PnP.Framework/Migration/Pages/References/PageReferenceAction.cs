using System.Collections.Generic;

namespace PnP.Framework.Migration.Pages.References
{
    public sealed class PageReferenceAction
    {
        public string SnapshotDependencyId { get; set; }

        public string TargetAbsoluteUrl { get; set; }

        public string TargetServerRelativeUrl { get; set; }

        public PageReferenceDisposition Disposition { get; set; }

        public IList<string> Diagnostics { get; set; } = new List<string>();
    }
}
