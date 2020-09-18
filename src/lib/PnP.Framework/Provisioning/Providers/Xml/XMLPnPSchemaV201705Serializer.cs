namespace PnP.Framework.Provisioning.Providers.Xml
{
    /// <summary>
    /// Implements the logic to serialize a schema of version 201705
    /// </summary>
    internal class XMLPnPSchemaV201705Serializer : XmlPnPSchemaBaseSerializer<V201705.ProvisioningTemplate>
    {
        public XMLPnPSchemaV201705Serializer() :
            base(typeof(XMLConstants)
                .Assembly
                .GetManifestResourceStream("PnP.Framework.Provisioning.Providers.Xml.ProvisioningSchema-2017-05.xsd"))
        {
        }

        public override string NamespaceUri
        {
            get
            {
                return (
#pragma warning disable 0618
                    XMLConstants.PROVISIONING_SCHEMA_NAMESPACE_2017_05
#pragma warning restore 0618
                    );
            }
        }

        public override string NamespacePrefix
        {
            get { return (XMLConstants.PROVISIONING_SCHEMA_PREFIX); }
        }

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

