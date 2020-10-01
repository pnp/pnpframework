using System.Text.Json.Serialization;

namespace PnP.Framework.Provisioning.Model.Configuration.Lists
{
    public class ApplyListsConfiguration
    {
        [JsonPropertyName("ignoreDuplicateDataRowErrors")]
        public bool IgnoreDuplicateDataRowErrors { get; set; }
    }
}
