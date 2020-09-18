using PnP.Framework.Provisioning.Model;
using PnP.Framework.Utilities;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Xml;
using System.Xml.Linq;
using System.Xml.Schema;
using System.Xml.Serialization;
using PnP.Framework.Provisioning.Providers.Xml.V201605;
using ContentType = PnP.Framework.Provisioning.Model.ContentType;
using PnP.Framework.Extensions;
using Microsoft.SharePoint.Client;
using Newtonsoft.Json.Serialization;
using PnP.Framework.Provisioning.Providers.Xml.Serializers;
using FileLevel = PnP.Framework.Provisioning.Model.FileLevel;

namespace PnP.Framework.Provisioning.Providers.Xml
{
    /// <summary>
    /// Implements the logic to serialize a schema of version 201605
    /// </summary>
    internal class XMLPnPSchemaV201605Serializer : XmlPnPSchemaBaseSerializer<V201605.ProvisioningTemplate>
    {
        public XMLPnPSchemaV201605Serializer():
            base(typeof(XMLConstants)
                .Assembly
                .GetManifestResourceStream("OfficeDevPnP.Core.Framework.Provisioning.Providers.Xml.ProvisioningSchema-2016-05.xsd"))
        {
        }

        public override string NamespaceUri
        {
            get { return (
#pragma warning disable 0618
                    XMLConstants.PROVISIONING_SCHEMA_NAMESPACE_2016_05
#pragma warning restore 0618
                    ); }
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

