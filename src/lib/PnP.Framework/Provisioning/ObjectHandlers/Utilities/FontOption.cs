using PnP.Core.Model.SharePoint;
using System.Text.Json.Serialization;

namespace PnP.Framework.Provisioning.ObjectHandlers.Utilities
{
    internal class FontOption : IFontOption
    {
        /// <summary>
        /// fontFamilyKey
        /// </summary>
        [JsonPropertyName("fontFamilyKey")]
        public string FamilyKey { get; set; }
        /// <summary>
        /// fontFace
        /// </summary>
        [JsonPropertyName("fontFace")]
        public string Face { get; set; }
        /// <summary>
        /// fontVariantWeight
        /// </summary>
        [JsonPropertyName("fontVariantWeight")]
        public string VariantWeight { get; set; }
    }
}
