using System.Collections.Generic;
using System.Text.Json.Serialization;

namespace PnP.Framework.Provisioning.Model.Configuration.ContentTypes
{
    public class ExtractContentTypeConfiguration
    {
        [JsonPropertyName("groups")]
        public List<string> Groups { get; set; } = new List<string>();

        [JsonPropertyName("excludeFromSyndication")]
        public bool ExcludeFromSyndication { get; set; }
    }
}
