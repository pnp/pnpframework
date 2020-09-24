using System;

namespace PnP.Framework.Entities
{
    /// <summary>
    /// Defines a Unified Group
    /// </summary>
    public class UnifiedGroupEntity
    {
        /// <summary>
        /// Unified group id
        /// </summary>
        public string GroupId { get; set; }
        /// <summary>
        /// Unified group display name
        /// </summary>
        public string DisplayName { get; set; }
        /// <summary>
        /// Unified group description 
        /// </summary>
        public string Description { get; set; }
        /// <summary>
        /// Unified group mail
        /// </summary>
        public string Mail { get; set; }
        /// <summary>
        /// Unified group nick name
        /// </summary>
        public string MailNickname { get; set; }
        /// <summary>
        /// Url of site to configure unified group
        /// </summary>
        public string SiteUrl { get; set; }
        /// <summary>
        /// Classification of the Office 365 group
        /// </summary>
        public string Classification { get; set; }
        /// <summary>
        /// Visibility of the Office 365 group
        /// </summary>
        public string Visibility { get; set; }
        /// <summary>
        /// Indication if the Office 365 Group has a Microsoft Team provisioned for it
        /// </summary>
        public bool? HasTeam { get; set; }
    }
}
