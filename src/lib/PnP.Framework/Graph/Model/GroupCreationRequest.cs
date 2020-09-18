using Newtonsoft.Json;

namespace PnP.Framework.Graph.Model
{
    public class GroupCreationRequest
    {
        [JsonProperty("displayName")]
        public string DisplayName { get; set; }

        [JsonProperty("description")]
        public string Description { get; set; }

        [JsonProperty("groupTypes")]
        public string[] GroupTypes { get; set; } = new string[] { "Unified" };

        [JsonProperty("mailEnabled")]
        public bool MailEnabled { get; set; } = true;

        [JsonProperty("securityEnabled")]
        public bool SecurityEnabled { get; set; } = false;

        [JsonProperty("mailNickname")]
        public string MailNickname { get; set; }

        [JsonProperty("owners@odata.bind")]
        public string[] Owners { get; set; }

        [JsonProperty("members@odata.bind")]
        public string[] Members { get; set; }
    }
}
