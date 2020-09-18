using Newtonsoft.Json;

namespace PnP.Framework.Provisioning.Model.Configuration.Lists
{
    public class ApplyListsConfiguration
    {
        [JsonProperty("ignoreDuplicateDataRowErrors")]
        public bool IgnoreDuplicateDataRowErrors { get; set; }
    }
}
