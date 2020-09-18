using Newtonsoft.Json;

namespace PnP.Framework.Provisioning.Model.Configuration.SiteSecurity
{
    public class ExtractConfiguration
    {
        [JsonProperty("includeSiteGroups")]
        public bool IncludeSiteGroups { get; set; }
    }
}
