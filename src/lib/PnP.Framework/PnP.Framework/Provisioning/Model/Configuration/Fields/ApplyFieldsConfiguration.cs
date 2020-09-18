using Newtonsoft.Json;

namespace PnP.Framework.Provisioning.Model.Configuration.Fields
{
    public class ApplyFieldsConfiguration
    {
        [JsonProperty("provisionFieldsToSubWebs")]
        public bool ProvisionFieldsToSubWebs { get; set; }
    }
}
