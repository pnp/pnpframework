using PnP.Framework.Migration.Evidence;

namespace PnP.Framework.Migration.Packaging
{
    public sealed class ArtifactReference
    {
        public string Sha256 { get; set; }

        public long Length { get; set; }

        public string MediaType { get; set; }

        public string ContentEncoding { get; set; }

        public string OriginalName { get; set; }

        public EvidenceAvailability Availability { get; set; } = EvidenceAvailability.Captured;

        public ArtifactLineage Lineage { get; set; }
    }
}
