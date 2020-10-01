using System.Text.Json.Serialization;

namespace PnP.Framework.Provisioning.Model.Configuration.SiteSecurity
{
    public class ExtractConfiguration
    {
        [JsonPropertyName("includeSiteGroups")]
        public bool IncludeSiteGroups { get; set; }
    }
}
