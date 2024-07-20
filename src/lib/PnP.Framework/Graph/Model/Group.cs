using PnP.Framework.Entities;

namespace PnP.Framework.Graph.Model
{
    internal class Group
    {
        public string GroupId { get; set; }
        public string DisplayName { get; set; }
        public string Description { get; set; }
        public string MailNickname { get; set; }
        /// <summary>
        /// Group e-mail address
        /// </summary>
        public string Mail { get; set; }
        public bool? MailEnabled { get; set; }
        /// <summary>
        /// Can the group be used to set permissions
        /// </summary>
        public bool? SecurityEnabled { get; set; }
        public string[] GroupTypes { get; set; }
        public string Visibility { get; set; }
        public string Classification { get; set; }

        public GroupEntity AsEntity()
        {
            return new GroupEntity()
            {
                Description = Description,
                DisplayName = DisplayName,
                GroupId = GroupId,
                GroupTypes = GroupTypes,
                MailNickname = MailNickname,
                Mail = Mail,
                MailEnabled = MailEnabled,
                SecurityEnabled = SecurityEnabled,
            };
        }
    }
}
