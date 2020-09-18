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
    /// Implements the logic to serialize a schema of version 201807
    /// </summary>
    internal class XMLPnPSchemaV201807Serializer : XmlPnPSchemaBaseSerializer<V201807.ProvisioningTemplate>
    {
        public XMLPnPSchemaV201807Serializer():
            base(typeof(XMLConstants)
                .Assembly
                .GetManifestResourceStream("OfficeDevPnP.Core.Framework.Provisioning.Providers.Xml.ProvisioningSchema-2018-07.xsd"))
        {
        }

        public override string NamespaceUri
        {
            get { return (XMLConstants.PROVISIONING_SCHEMA_NAMESPACE_2018_07); }
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

