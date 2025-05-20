using System;
using System.Text.Json.Serialization;

namespace PnP.Framework.Entities
{
    /// <summary>
    /// Defines a Unified\Microsoft 365 Group user
    /// </summary>
    public class UnifiedGroupUser
    {
        /// <summary>
        /// Group user's unique identifier
        /// </summary>
        public string Id { get; set; }        
        /// <summary>
        /// Group user's user principal name
        /// </summary>
        public string UserPrincipalName { get; set; }
        /// <summary>
        /// Group user's display name
        /// </summary>
        public string DisplayName { get; set; }
        /// <summary>
        /// Group user's given name
        /// </summary>
        public string GivenName { get; set; }
        /// <summary>
        /// Group user's surname
        /// </summary>
        public string Surname { get; set; }
        /// <summary>
        /// Group user's e-mail address
        /// </summary>
        [JsonPropertyName("mail")]
        public string Email { get; set; }
        /// <summary>
        /// Group user's mobile phone number
        /// </summary>
        public string MobilePhone { get; set; }        
        /// <summary>
        /// Group user's preferred language in ISO 639-1 standard notation
        /// </summary>
        public string PreferredLanguage { get; set; }
        /// <summary>
        /// Group user's job title
        /// </summary>
        public string JobTitle { get; set; }
        /// <summary>
        /// Group user's business phone numbers
        /// </summary>
        public string[] BusinessPhones { get; set; }        
    }
}
