using System.Collections.Generic;

namespace PnP.Framework.Migration.Pages.Runtime
{
    public sealed class PageRuntimeSnapshot
    {
        public string SchemaVersion { get; set; } = "pnp-page-runtime/v1";

        public string PageDeclaredType { get; set; }

        public string LayoutDeclaredType { get; set; }

        public string AdapterId { get; set; } = PageRuntimeAdapterIds.Unknown;

        public PageRuntimeDetectionSource DetectionSource { get; set; }

        public PageRuntimeResolutionState ResolutionState { get; set; }

        public IList<string> Diagnostics { get; set; } = new List<string>();
    }
}
