using System.Text.Json.Serialization;

namespace PnP.Framework.Provisioning.Model.Configuration.ContentTypes
{
    public class ApplyContentTypeConfiguration
    {
        [JsonPropertyName("provisionContentTypesToSubWebs")]
        public bool ProvisionContentTypesToSubWebs { get; set; }
    }
}
