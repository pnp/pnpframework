using PnP.Framework.Migration.Evidence;
using PnP.Framework.Migration.Packaging;
using System;
using System.Collections.Generic;

namespace PnP.Framework.Migration.Pages.Markup
{
    public sealed class PageArtifactSnapshot
    {
        public string SchemaVersion { get; set; } = "pnp-page-artifact/v1";

        public Guid FileUniqueId { get; set; }

        public string ServerRelativeUrl { get; set; }

        public ArtifactReference Bytes { get; set; }

        public string ContentBase64 { get; set; }

        public PageDirectiveSnapshot PageDirective { get; set; }

        public EvidenceAvailability Availability { get; set; }

        public IList<string> Diagnostics { get; set; } = new List<string>();
    }
}
