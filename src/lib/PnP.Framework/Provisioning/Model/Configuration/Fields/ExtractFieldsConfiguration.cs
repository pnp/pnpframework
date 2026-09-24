using System.Collections.Generic;
using System.Text.Json.Serialization;

namespace PnP.Framework.Provisioning.Model.Configuration.Fields
{
    public class ExtractFieldsConfiguration
    {
        [JsonPropertyName("groups")]
        public List<string> Groups { get; set; } = new List<string>();
    }
}
