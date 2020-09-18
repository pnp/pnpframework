using Newtonsoft.Json;
using System.Collections.Generic;

namespace PnP.Framework.Provisioning.Model.Configuration.Tenant.Teams
{
    public class ExtractTeamsConfiguration
    {
        [JsonProperty("includeAllTeams")]
        public bool IncludeAllTeams { get; set; }

        [JsonProperty("includeMessages")]
        public bool IncludeMessages { get; set; }

        [JsonProperty("teamSiteUrls")]
        public List<string> TeamSiteUrls { get; set; } = new List<string>();

        [JsonProperty("includeGroupId")]
        public bool IncludeGroupId { get; set; }
    }
}
