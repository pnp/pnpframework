namespace PnP.Framework.Provisioning.Providers.Xml
{
    /// <summary>
    /// Implements the logic to serialize a schema of version 201801
    /// </summary>
    internal class XMLPnPSchemaV201801Serializer : XmlPnPSchemaBaseSerializer<V201801.ProvisioningTemplate>
    {
        public XMLPnPSchemaV201801Serializer() :
            base(typeof(XMLConstants)
                .Assembly
                .GetManifestResourceStream("PnP.Framework.Provisioning.Providers.Xml.ProvisioningSchema-2018-01.xsd"))
        {
        }

#pragma warning disable CS0618
        public override string NamespaceUri => XMLConstants.PROVISIONING_SCHEMA_NAMESPACE_2018_01;
#pragma warning restore CS0618

        public override string NamespacePrefix => XMLConstants.PROVISIONING_SCHEMA_PREFIX;

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

