using System.Text.Json.Serialization;

namespace PnP.Framework.Provisioning.Model.Configuration.Publishing
{
    public class ExtractPublishingConfiguration
    {
        [JsonPropertyName("includeNativePublishingFiles")]
        public bool IncludeNativePublishingFiles { get; set; }

        [JsonPropertyName("persist")]
        public bool Persist { get; set; }
    }
}
