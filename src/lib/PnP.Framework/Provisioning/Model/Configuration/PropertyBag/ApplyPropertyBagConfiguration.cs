using System.Text.Json.Serialization;

namespace PnP.Framework.Provisioning.Model.Configuration.PropertyBag
{
    public class ApplyPropertyBagConfiguration
    {
        [JsonPropertyName("overwriteSystemValues")]
        public bool OverwriteSystemValues { get; set; }
    }
}
