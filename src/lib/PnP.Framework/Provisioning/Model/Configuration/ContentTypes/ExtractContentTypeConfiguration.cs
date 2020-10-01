using System.Collections.Generic;
using System.Text.Json.Serialization;

namespace PnP.Framework.Provisioning.Model.Configuration.ContentTypes
{
    public class ExtractContentTypeConfiguration
    {
        [JsonPropertyName("Groups")]
        public List<string> Groups { get; set; } = new List<string>();

        [JsonPropertyName("IncludeFromSyndication")]
        public bool ExcludeFromSyndication { get; set; }
    }
}
