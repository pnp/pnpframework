using System;

namespace PnP.Framework.Modernization.Entities
{
    /// <summary>
    /// Class to hold a mapping between a user in the source site and a user in the target site
    /// </summary>
    [Serializable]
    public class UserMappingEntity
    {
        /// <summary>
        /// Source user reference
        /// </summary>
        public string SourceUser { get; set; }

        /// <summary>
        /// Target user reference
        /// </summary>
        public string TargetUser { get; set; }
    }
}
