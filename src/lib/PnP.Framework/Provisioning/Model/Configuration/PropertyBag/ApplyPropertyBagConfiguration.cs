using Newtonsoft.Json;

namespace PnP.Framework.Provisioning.Model.Configuration.PropertyBag
{
    public class ApplyPropertyBagConfiguration
    {
        [JsonProperty("overwriteSystemValues")]
        public bool OverwriteSystemValues { get; set; }
    }
}
