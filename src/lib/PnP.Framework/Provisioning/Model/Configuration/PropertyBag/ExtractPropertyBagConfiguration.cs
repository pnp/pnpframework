using Newtonsoft.Json;
using System.Collections.Generic;

namespace PnP.Framework.Provisioning.Model.Configuration.PropertyBag
{
    public class ExtractPropertyBagConfiguration
    {
        [JsonProperty("valuesToPreserve")]
        internal List<string> ValuesToPreserve { get; set; }
    }
}
