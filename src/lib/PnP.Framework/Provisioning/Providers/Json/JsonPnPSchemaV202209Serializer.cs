using Newtonsoft.Json;

namespace PnP.Framework.Provisioning.Providers.Json
{
    internal class JsonPnPSchemaV202209Serializer : JsonPnPSchemaBaseSerializer<Xml.V202209.ProvisioningTemplate>
    {

        public JsonPnPSchemaV202209Serializer() 
            : this(null) { }

        public JsonPnPSchemaV202209Serializer(JsonSerializerSettings serializerSettings)
            : base(serializerSettings) { }

        protected override void DeserializeTemplate(object persistenceTemplate, Model.ProvisioningTemplate template)
        {
            base.DeserializeTemplate(persistenceTemplate, template);
        }

        protected override void SerializeTemplate(Model.ProvisioningTemplate template, object persistenceTemplate)
        {
            base.SerializeTemplate(template, persistenceTemplate);
        }
    }
}
