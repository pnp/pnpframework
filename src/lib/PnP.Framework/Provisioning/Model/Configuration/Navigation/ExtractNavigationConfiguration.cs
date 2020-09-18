using Newtonsoft.Json;

namespace PnP.Framework.Provisioning.Model.Configuration.Navigation
{
    public class ExtractNavigationConfiguration
    {
        [JsonProperty("RemoveExistingNodes")]
        public bool RemoveExistingNodes { get; set; }
    }
}
