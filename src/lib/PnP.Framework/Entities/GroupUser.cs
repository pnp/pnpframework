using PnP.Framework.Enums;

namespace PnP.Framework.Entities
{
    /// <summary>
    /// Defines an user or group located in a Azure Active Directory Group
    /// </summary>
    public class GroupUser
    {
        /// <summary>
        /// Group user's user principal name or Group Id
        /// </summary>
        public string UserPrincipalName { get; set; }

        /// <summary>
        /// Group user's or group's display name
        /// </summary>
        public string DisplayName { get; set; }

        /// <summary>
        /// Indication if this entry represents a user or a group
        /// </summary>
        public GroupUserType Type { get; set; }
    }
}
