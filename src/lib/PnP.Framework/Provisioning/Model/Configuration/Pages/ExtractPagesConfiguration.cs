using System.Text.Json.Serialization;

namespace PnP.Framework.Provisioning.Model.Configuration.Pages
{
    public class ExtractPagesConfiguration
    {
        [JsonPropertyName("excludeAuthorInformation")]
        public bool ExcludeAuthorInformation { get; set; }

        [JsonPropertyName("includeAllClientSidePages")]
        public bool IncludeAllClientSidePages { get; set; }
    }
}
