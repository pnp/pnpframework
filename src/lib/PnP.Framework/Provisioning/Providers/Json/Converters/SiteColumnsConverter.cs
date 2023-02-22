using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using PnP.Framework.Extensions;
using PnP.Framework.Provisioning.Providers.Xml;
using System;

namespace PnP.Framework.Provisioning.Providers.Json.Converters
{
    public class SiteColumnsConverter : JsonConverter
    {
        public override bool CanConvert(Type objectType)
        {
            var fieldsTypeName = $"{PnPSerializationScope.Current?.BaseSchemaNamespace}.ProvisioningTemplateSiteFields, {PnPSerializationScope.Current?.BaseSchemaAssemblyName}";
            var fieldsType = Type.GetType(fieldsTypeName, true);
            return fieldsType.IsAssignableFrom(objectType);
        }

        public override object ReadJson(JsonReader reader, Type objectType, object existingValue, JsonSerializer serializer)
        {
            JArray sourceJObject = JArray.Load(reader);
            var e = sourceJObject.ToObject<System.Xml.XmlElement[]>();

            var fields = Activator.CreateInstance(objectType);
            fields.SetPublicInstancePropertyValue("Any", e);

            return fields;
        }

        public override void WriteJson(JsonWriter writer, object value, JsonSerializer serializer)
        {
            JToken t = JToken.FromObject(value);

            if (t.Type != JTokenType.Object)
            {
                t.WriteTo(writer);
            }
            else
            {
                JArray a = (JArray)t.SelectToken("Any");
                a.WriteTo(writer);
            }
        }
    }
}
