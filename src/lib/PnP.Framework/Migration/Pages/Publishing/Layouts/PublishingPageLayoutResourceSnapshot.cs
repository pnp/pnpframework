using PnP.Framework.Migration.Evidence;
using PnP.Framework.Migration.Packaging;
using System.Collections.Generic;

namespace PnP.Framework.Migration.Pages.Publishing.Layouts
{
    public sealed class PublishingPageLayoutResourceSnapshot
    {
        public PublishingPageLayoutResourceReference Reference { get; set; }

        public PublishingPageLayoutResourceEvidenceState EvidenceState { get; set; }

        public string ResolvedSourceUrl { get; set; }

        public ArtifactReference Artifact { get; set; }

        public string ContentBase64 { get; set; }

        public IList<EvidenceSource> Sources { get; set; } = new List<EvidenceSource>();

        public IList<string> Diagnostics { get; set; } = new List<string>();
    }
}
