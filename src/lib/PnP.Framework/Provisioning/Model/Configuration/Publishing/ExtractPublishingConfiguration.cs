using Newtonsoft.Json;

namespace PnP.Framework.Provisioning.Model.Configuration.Publishing
{
    public class ExtractPublishingConfiguration
    {
        [JsonProperty("includeNativePublishingFiles")]
        public bool IncludeNativePublishingFiles { get; set; }

        [JsonProperty("persist")]
        public bool Persist { get; set; }
    }
}
