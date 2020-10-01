using System.Text.Json.Serialization;

namespace PnP.Framework.Provisioning.Model.Configuration.SiteFooter
{
    public class ExtractSiteFooterConfiguration
    {
        [JsonPropertyName("removeExistingNodes")]
        public bool RemoveExistingNodes { get; set; }
    }
}
