using Newtonsoft.Json;

namespace PnP.Framework.Provisioning.Model.Configuration.ContentTypes
{
    public class ApplyContentTypeConfiguration
    {
        [JsonProperty("provisionContentTypesToSubWebs")]
        public bool ProvisionContentTypesToSubWebs { get; set; }
    }
}
