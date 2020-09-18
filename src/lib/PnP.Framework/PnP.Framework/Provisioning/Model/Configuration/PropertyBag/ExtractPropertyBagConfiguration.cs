using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PnP.Framework.Provisioning.Model.Configuration.PropertyBag
{
    public class ExtractPropertyBagConfiguration
    {
        [JsonProperty("valuesToPreserve")]
        internal List<string> ValuesToPreserve { get; set; }
    }
}
