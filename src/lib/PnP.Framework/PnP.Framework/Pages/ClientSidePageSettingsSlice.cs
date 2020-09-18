using Newtonsoft.Json;

namespace PnP.Framework.Pages
{
    public class ClientSidePageSettingsSlice
    {
        [JsonProperty(PropertyName = "isDefaultDescription", NullValueHandling = NullValueHandling.Ignore)]
        public bool? IsDefaultDescription { get; set; }

        [JsonProperty(PropertyName = "isDefaultThumbnail", NullValueHandling = NullValueHandling.Ignore)]
        public bool? IsDefaultThumbnail { get; set; }
    }
}
