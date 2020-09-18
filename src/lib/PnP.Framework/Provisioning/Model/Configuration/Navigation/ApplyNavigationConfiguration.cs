using Newtonsoft.Json;

namespace PnP.Framework.Provisioning.Model.Configuration.Navigation
{
    public class ApplyNavigationConfiguration
    {
        [JsonProperty("clearNavigation")]
        public bool ClearNavigation { get; set; }
    }
}
