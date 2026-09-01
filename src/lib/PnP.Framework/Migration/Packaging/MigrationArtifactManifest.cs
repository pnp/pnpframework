using System.Collections.Generic;

namespace PnP.Framework.Migration.Packaging
{
    public sealed class MigrationArtifactManifest
    {
        public string SchemaVersion { get; set; } = "pnp-migration-artifacts/v1";

        public IList<ArtifactReference> Artifacts { get; set; } = new List<ArtifactReference>();

        public string ContentSha256 { get; set; }
    }
}
