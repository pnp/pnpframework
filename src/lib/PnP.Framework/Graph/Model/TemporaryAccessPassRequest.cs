using System;
using Newtonsoft.Json;

namespace PnP.Framework.Graph.Model
{
    /// <summary>
    /// Defines a request for a temporary access pass for a User
    /// </summary>
    public class TemporaryAccessPassRequest
    {
        /// <summary>
        /// Indicates the type(s) of change(s) in the subscribed resource that will raise a notification
        /// </summary>
        [JsonProperty("@odata.type")]
        public string ODataType => "#microsoft.graph.temporaryAccessPassAuthenticationMethod";

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
    }
}
