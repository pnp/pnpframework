using System.Text.Json.Serialization;

namespace PnP.Framework.Provisioning.Model.Configuration.MultiLanguage
{
    public class ExtractMultiLanguageConfiguration
    {
        [JsonPropertyName("persistMultilanguageResources")]
        public bool PersistResources { get; set; }

        [JsonPropertyName("resourceFilePrefix")]
        public string ResourceFilePrefix { get; set; }
    }
}
