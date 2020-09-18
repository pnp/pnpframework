using Newtonsoft.Json;

namespace PnP.Framework.Provisioning.Model.Configuration.SearchSettings
{
    public class ExtractSearchConfiguration
    {
        [JsonProperty("include")]
        public bool Include { get; set; }
    }
}
