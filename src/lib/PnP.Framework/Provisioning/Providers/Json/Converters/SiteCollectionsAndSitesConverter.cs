using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using PnP.Framework.Provisioning.Model;
using PnP.Framework.Provisioning.Providers.Xml;
using System;

namespace PnP.Framework.Provisioning.Providers.Json.Converters
{
    public class SiteCollectionsAndSitesConverter : JsonConverter
    {
        public override bool CanConvert(Type objectType)
        {
            var siteTypeName = $"{PnPSerializationScope.Current?.BaseSchemaNamespace}.SiteCollection, {PnPSerializationScope.Current?.BaseSchemaAssemblyName}";
            var siteType = Type.GetType(siteTypeName, true);
            return siteType.IsAssignableFrom(objectType);
        }

        public override object ReadJson(JsonReader reader, Type objectType, object existingValue, JsonSerializer serializer)
        {
            // Define the specific source schema types
            var communicationSiteTypeName = $"{PnPSerializationScope.Current?.BaseSchemaNamespace}.CommunicationSite, {PnPSerializationScope.Current?.BaseSchemaAssemblyName}";
            var communicationSiteType = Type.GetType(communicationSiteTypeName, true);
            var teamSiteTypeName = $"{PnPSerializationScope.Current?.BaseSchemaNamespace}.TeamSite, {PnPSerializationScope.Current?.BaseSchemaAssemblyName}";
            var teamSiteType = Type.GetType(teamSiteTypeName, true);
            var teamSiteNoGroupTypeName = $"{PnPSerializationScope.Current?.BaseSchemaNamespace}.TeamSiteNoGroup, {PnPSerializationScope.Current?.BaseSchemaAssemblyName}";
            var teamSiteNoGroupType = Type.GetType(teamSiteNoGroupTypeName, true);
            var teamSubSiteNoGroupTypeName = $"{PnPSerializationScope.Current?.BaseSchemaNamespace}.TeamSubSiteNoGroup, {PnPSerializationScope.Current?.BaseSchemaAssemblyName}";
            var teamSubSiteNoGroupType = Type.GetType(teamSubSiteNoGroupTypeName, true);

            JObject sourceJObject = JObject.Load(reader);
            JToken discriminatorToken = sourceJObject.SelectToken("Type");
            string discriminator = discriminatorToken?.Value<string>() ?? "CommunicationSite";
            discriminatorToken.Parent.Remove();

            Object targetItem = null;
            switch (discriminator)
            {
                case "CommunicationSite":
                    targetItem = sourceJObject.ToObject(communicationSiteType);
                    break;
                case "TeamSite":
                    targetItem = sourceJObject.ToObject(teamSiteType);
                    break;
                case "TeamSiteNoGroup":
                    targetItem = sourceJObject.ToObject(teamSiteNoGroupType);
                    break;
                case "TeamNoGroupSubSite":
                    targetItem = sourceJObject.ToObject(teamSubSiteNoGroupType);
                    break;
            }

            return targetItem;
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
                string typeProperty =value.GetType().Name;

                JObject o = (JObject)t;
                o.AddFirst(new JProperty("Type", typeProperty));

                o.WriteTo(writer);
            }
        }
    }
}
