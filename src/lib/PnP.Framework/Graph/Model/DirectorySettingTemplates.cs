using Newtonsoft.Json;
using System.Collections.Generic;

namespace PnP.Framework.Graph.Model
{
    /// <summary>
    /// Defines the container for a collection of DirectorySetting objects
    /// </summary>
    public class DirectorySettingTemplates
    {
        [JsonProperty(PropertyName = "value")]
        public List<DirectorySetting> Templates { get; set; }
    }
}
