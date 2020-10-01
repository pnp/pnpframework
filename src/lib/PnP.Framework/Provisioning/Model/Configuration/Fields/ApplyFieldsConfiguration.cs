using System.Text.Json.Serialization;

namespace PnP.Framework.Provisioning.Model.Configuration.Fields
{
    public class ApplyFieldsConfiguration
    {
        [JsonPropertyName("provisionFieldsToSubWebs")]
        public bool ProvisionFieldsToSubWebs { get; set; }
    }
}
