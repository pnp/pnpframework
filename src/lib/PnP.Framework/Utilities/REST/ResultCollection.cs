using System.Collections.Generic;
using System.Text.Json.Serialization;

namespace PnP.Framework.Utilities.REST
{
    public class ResultCollection<T>
    {
        [JsonPropertyName("@odata.nextLink")]
        public string NextLink { get; set; }

        [JsonPropertyName("value")]
        public IEnumerable<T> Items { get; set; }
    }
}
