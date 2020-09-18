using Newtonsoft.Json;
using PnP.Framework.Utilities.JsonConverters;

namespace PnP.Framework.Pages
{
    public class ClientSideSectionEmphasis
    {
        [JsonProperty(PropertyName = "zoneEmphasis", NullValueHandling = NullValueHandling.Ignore)]
        [JsonConverter(typeof(EmphasisJsonConverter))]
        public int ZoneEmphasis
        {
            get; set;
        }
    }
}
