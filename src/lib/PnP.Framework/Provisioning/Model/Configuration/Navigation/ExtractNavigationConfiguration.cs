using System.Text.Json.Serialization;

namespace PnP.Framework.Provisioning.Model.Configuration.Navigation
{
    public class ExtractNavigationConfiguration
    {
        [JsonPropertyName("RemoveExistingNodes")]
        public bool RemoveExistingNodes { get; set; }
    }
}
