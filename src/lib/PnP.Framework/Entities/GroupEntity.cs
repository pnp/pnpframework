using System;

namespace PnP.Framework.Entities
{
    /// <summary>
    /// Defines an Azure Active Directory Group
    /// </summary>
    public class GroupEntity
    {
        /// <summary>
        /// Group id
        /// </summary>
        public string GroupId { get; set; }
        /// <summary>
        /// Group display name
        /// </summary>
        public string DisplayName { get; set; }
        /// <summary>
        /// Group description 
        /// </summary>
        public string Description { get; set; }
        /// <summary>
        /// Group nick name
        /// </summary>
        public string MailNickname { get; set; }
        /// <summary>
        /// Group e-mail address
        /// </summary>
        public string Mail { get; set; }        
        /// <summary>
        /// Is the group enabled for receiving e-mail
        /// </summary>
        public bool? MailEnabled { get; set; }
        /// <summary>
        /// Can the group be used to set permissions
        /// </summary>
        public bool? SecurityEnabled { get; set; }     
        /// <summary>
        /// Types of group
        /// </summary>
        public string[] GroupTypes { get; set; }             
    }
}
