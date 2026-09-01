using System.Collections.Generic;

namespace PnP.Framework.Migration.Evidence
{
    public sealed class ArtifactLineage
    {
        public IList<string> InputExchangeIds { get; set; } = new List<string>();

        public IList<string> InputPayloadSha256 { get; set; } = new List<string>();

        public string ProjectorId { get; set; }

        public string ProjectorVersion { get; set; }

        public string OutputSchemaVersion { get; set; }

        public string OutputSha256 { get; set; }
    }
}
