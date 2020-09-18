using Newtonsoft.Json;

namespace PnP.Framework.Pages
{
    /// <summary>
    /// Control data for controls of type 3 (= client side web parts) which persist using the data-sp-controldata property only
    /// </summary>
    public class ClientSideWebPartControlDataOnly : ClientSideWebPartControlData
    {
        [JsonProperty(PropertyName = "webPartData", NullValueHandling = NullValueHandling.Ignore)]
        public string WebPartData { get; set; }
    }
}
