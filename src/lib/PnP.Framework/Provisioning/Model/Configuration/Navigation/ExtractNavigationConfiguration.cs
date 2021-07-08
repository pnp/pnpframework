using System.Text.Json.Serialization;

namespace PnP.Framework.Provisioning.Model.Configuration.Navigation
{
    public class ExtractNavigationConfiguration
    {
        [JsonPropertyName("removeExistingNodes")]
        public bool RemoveExistingNodes { get; set; }
    }
}
