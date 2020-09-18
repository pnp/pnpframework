using Newtonsoft.Json;
using System.Collections.Generic;

namespace PnP.Framework.Provisioning.Model.Configuration.Tenant.Sequence
{
    public class ExtractSequenceConfiguration
    {
        [JsonProperty("siteUrls")]
        public List<string> SiteUrls { get; set; } = new List<string>();

        [JsonProperty("maxSubsiteDepth")]
        public int MaxSubsiteDepth { get; set; }

        [JsonProperty("includeJoinedSites")]
        public bool IncludeJoinedSites { get; set; }

        [JsonProperty("includeSubsites")]
        public bool IncludeSubsites { get; set; }
    }
}
