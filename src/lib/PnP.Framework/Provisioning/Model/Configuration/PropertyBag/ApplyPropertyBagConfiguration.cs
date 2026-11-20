using System.Text.Json.Serialization;

namespace PnP.Framework.Provisioning.Model.Configuration.PropertyBag
{
    public class ApplyPropertyBagConfiguration
    {
        [JsonPropertyName("overwriteSystemValues")]
        public bool OverwriteSystemValues { get; set; }

        /// <summary>
        /// When null, checks property bag write access on the target web.
        /// Set to true to attempt writes without checking, or false to skip writes.
        /// </summary>
        [JsonPropertyName("propertyBagWriteAllowed")]
        public bool? PropertyBagWriteAllowed { get; set; }
    }
}
