using System;
using Newtonsoft.Json;

namespace PnP.Framework.Graph.Model
{
    /// <summary>
    /// Defines a response for a temporary access pass for a User
    /// </summary>
    public class TemporaryAccessPassResponse
    {
        /// <summary>
        /// Identifier of the temporary access pass
        /// </summary>
        [JsonProperty("id")]
        public Guid? Id { get; set; }  

        /// <summary>
        /// The temporary access pass code
        /// </summary>
        [JsonProperty("temporaryAccessPass")]
        public string TemporaryAccessPass { get; set; }  

        /// <summary>
        /// Date and time on which the temporary access pass has been created
        /// </summary>
        [JsonProperty("createdDateTime")]
        public DateTime? CreatedDateTime { get; set; }

        /// <summary>
        /// Date and time on which the temporary access pass should become valid. If not provided, the access pass will be valid immediately.
        /// </summary>
        [JsonProperty("startDateTime")]
        public DateTime? StartDateTime { get; set; }

        /// <summary>
        /// The time in minutes specifying how long the temporary access pass should be valid for. If not provided, the default duration as configured in Azure Active Directory will be applied.
        /// </summary>
        [JsonProperty("lifetimeInMinutes")]
        public int? LifetimeInMinutes { get; set; }

        /// <summary>
        /// Boolean indicating if the temporary access pass can only be used once to log in (true) or continously for as long as the pass is valid for (false)
        /// </summary>
        [JsonProperty("isUsableOnce")]
        public bool? IsUsableOnce { get; set; }

        /// <summary>
        /// Boolean indicating if the temporary access pass can be used already
        /// </summary>
        [JsonProperty("isUsable")]
        public bool? IsUsable { get; set; }

        /// <summary>
        /// Provides more context around why the pass can or can not be used yet
        /// </summary>
        [JsonProperty("methodUsabilityReason")]
        public string MethodUsabilityReason { get; set; }             
    }
}
