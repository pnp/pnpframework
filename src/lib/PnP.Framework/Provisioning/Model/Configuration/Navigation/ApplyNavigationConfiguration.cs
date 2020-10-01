using System.Text.Json.Serialization;

namespace PnP.Framework.Provisioning.Model.Configuration.Navigation
{
    public class ApplyNavigationConfiguration
    {
        [JsonPropertyName("clearNavigation")]
        public bool ClearNavigation { get; set; }
    }
}
