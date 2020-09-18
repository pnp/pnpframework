using Newtonsoft.Json;
using System.Collections.Generic;

namespace PnP.Framework.Provisioning.Model.Configuration.Extensibility
{
    public class ApplyExtensibilityConfiguration
    {
        [JsonProperty("handlers")]
        public List<ExtensibilityHandler> Handlers { get; set; } = new List<ExtensibilityHandler>();
    }
}
