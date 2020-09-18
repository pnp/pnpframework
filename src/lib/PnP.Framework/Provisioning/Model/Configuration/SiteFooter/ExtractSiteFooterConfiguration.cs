using Newtonsoft.Json;

namespace PnP.Framework.Provisioning.Model.Configuration.SiteFooter
{
    public class ExtractSiteFooterConfiguration
    {
        [JsonProperty("removeExistingNodes")]
        public bool RemoveExistingNodes { get; set; }
    }
}
