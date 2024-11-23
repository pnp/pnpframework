using Newtonsoft.Json;
using System.Collections.Generic;

namespace PnP.Framework.Graph.Model
{
    internal class GroupExtended : Group
    {
#pragma warning disable CA1819
        [JsonProperty("owners@odata.bind", NullValueHandling = NullValueHandling.Ignore)]
        public string[] OwnersODataBind { get; set; }
        [JsonProperty("members@odata.bind", NullValueHandling = NullValueHandling.Ignore)]
        public string[] MembersODataBind { get; set; }
#pragma warning restore CA1819

        public List<GroupLabel> AssignedLabels { get; set; }
        public string PreferredDataLocation { get; set; }

        public Dictionary<string, object> AdditionalData { get; set; }
    }
}
