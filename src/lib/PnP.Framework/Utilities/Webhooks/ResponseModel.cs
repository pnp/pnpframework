using Newtonsoft.Json;
using System.Collections.Generic;

namespace PnP.Framework.Utilities.Webhooks
{

    /// <summary>
    /// 
    /// </summary>
    internal class ResponseModel<T>
    {

        [JsonProperty(PropertyName = "value")]
        public List<T> Value { get; set; }
    }
}
