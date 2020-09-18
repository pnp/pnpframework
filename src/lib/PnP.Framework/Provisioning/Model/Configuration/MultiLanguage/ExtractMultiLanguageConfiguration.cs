using Newtonsoft.Json;

namespace PnP.Framework.Provisioning.Model.Configuration.MultiLanguage
{
    public class ExtractMultiLanguageConfiguration
    {
        [JsonProperty("persistMultilanguageResources")]
        public bool PersistResources { get; set; }

        [JsonProperty("resourceFilePrefix")]
        public string ResourceFilePrefix { get; set; }
    }
}
