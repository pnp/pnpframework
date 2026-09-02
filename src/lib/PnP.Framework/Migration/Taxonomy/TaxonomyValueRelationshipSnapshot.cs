using System;
using System.Collections.Generic;

namespace PnP.Framework.Migration.Taxonomy
{
    public sealed class TaxonomyValueRelationshipSnapshot
    {
        public string SchemaVersion { get; set; } = "pnp-taxonomy-value-relationship/v1";

        public DateTimeOffset CapturedAtUtc { get; set; }

        public TaxonomyRelationshipState State { get; set; }

        public Guid? LiveTermSetId { get; set; }

        public string LiveTermSetName { get; set; }

        public string LiveTermLabel { get; set; }

        public string LiveTermPath { get; set; }

        public bool? LiveTermAvailableForTagging { get; set; }

        public TaxonomyHiddenListEntrySnapshot ValueHiddenListEntry { get; set; }

        public TaxonomyHiddenListEntrySnapshot TaxCatchAllHiddenListEntry { get; set; }

        public string SourceFieldValueSetSha256 { get; set; }

        public string EvidenceSha256 { get; set; }

        public IList<string> Diagnostics { get; set; } = new List<string>();
    }
}
