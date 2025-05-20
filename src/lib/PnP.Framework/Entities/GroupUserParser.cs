using PnP.Framework.Enums;
using System.Linq;
using System.Text.Json;
using System.Text.Json.Nodes;
using System.Text.Json.Serialization;

namespace PnP.Framework.Entities
{
    public class GroupUserParser
    {
        private class GroupUserRaw
        {

            public string Id { get; set; }
            public string UserPrincipalName { get; set; }

            public string DisplayName { get; set; }

            [JsonPropertyName("@odata.type")]
            public string Type { get; set; }
        }
        public static GroupUser[] ReadListFromJsonNode(JsonNode inputJson)
        {
            var rawList = inputJson.Deserialize<GroupUserRaw[]>();
            return rawList.Select(raw => raw.Type.Contains("microsoft.graph.user")
                ? new GroupUser
                {
                    DisplayName = raw.DisplayName,
                    UserPrincipalName = raw.UserPrincipalName,
                    Type = GroupUserType.User,
                }
                : new GroupUser
                {
                    DisplayName = raw.DisplayName,
                    UserPrincipalName = raw.Id,
                    Type = GroupUserType.Group,
                })
                .ToArray();
        }
    }
}
