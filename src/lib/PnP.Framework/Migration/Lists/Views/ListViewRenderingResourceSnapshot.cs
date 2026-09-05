using PnP.Framework.Migration.Evidence;
using PnP.Framework.Migration.Packaging;
using System.Collections.Generic;

namespace PnP.Framework.Migration.Lists.Views
{
    public enum ListViewRenderingResourceKind
    {
        JavaScript = 1,
        Xsl = 2,
        StyleSheet = 3,
        Other = 4
    }

    public sealed class ListViewRenderingResourceBindingSnapshot
    {
        public string SourceProperty { get; set; }

        public string OriginalReference { get; set; }

        public string ResourceId { get; set; }
    }

    public sealed class ListViewRenderingResourceSnapshot
    {
        public string Id { get; set; }

        public ListViewRenderingResourceKind Kind { get; set; }

        public string SourceAbsoluteUrl { get; set; }

        public string SourceServerRelativeUrl { get; set; }

        public ArtifactReference Artifact { get; set; }

        public string ContentBase64 { get; set; }

        public EvidenceAvailability Availability { get; set; } = EvidenceAvailability.Captured;

        public IList<string> Diagnostics { get; set; } = new List<string>();
    }
}
