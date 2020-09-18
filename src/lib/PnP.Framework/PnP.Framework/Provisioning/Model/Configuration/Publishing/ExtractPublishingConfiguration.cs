using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PnP.Framework.Provisioning.Model.Configuration.Publishing
{
    public class ExtractPublishingConfiguration
    {
        [JsonProperty("includeNativePublishingFiles")]
        public bool IncludeNativePublishingFiles { get; set; }

        [JsonProperty("persist")]
        public bool Persist { get; set; }
    }
}
