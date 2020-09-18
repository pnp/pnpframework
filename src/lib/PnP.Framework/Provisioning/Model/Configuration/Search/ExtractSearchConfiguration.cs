using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PnP.Framework.Provisioning.Model.Configuration.SearchSettings
{
    public class ExtractSearchConfiguration
    {
       [JsonProperty("include")]
       public bool Include { get; set; }
    }
}
