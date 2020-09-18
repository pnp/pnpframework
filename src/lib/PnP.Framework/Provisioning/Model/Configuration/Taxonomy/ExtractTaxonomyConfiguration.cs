using Newtonsoft.Json;

namespace PnP.Framework.Provisioning.Model.Configuration.Taxonomy
{
    public class ExtractTaxonomyConfiguration
    {
        [JsonProperty("includeSecurity")]
        public bool IncludeSecurity { get; set; }

        [JsonProperty("includeSiteCollectionTermGroup")]
        public bool IncludeSiteCollectionTermGroup { get; set; }

        [JsonProperty("includeAllTermGroups")]
        public bool IncludeAllTermGroups { get; set; }
    }
}
