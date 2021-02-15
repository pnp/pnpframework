using System;
using System.Collections.Generic;
using System.Text;
using System.Xml.Serialization;

namespace PnP.Framework.Utilities.UnitTests.Model
{
    [XmlRoot(Namespace = "http://schemas.microsoft.com/sharepoint/clientquery/2009", ElementName = "Request")]
    public class CSOMRequest
    {
        [XmlAttribute]
        public bool AddExpandoFieldTypeSuffix { get; set; }
        [XmlArrayItem(Type = typeof(BaseAction), ElementName = "ObjectPath"),
            XmlArrayItem(Type = typeof(QueryAction), ElementName = "Query"),
            XmlArrayItem(Type = typeof(MethodAction), ElementName = "Method"),
            XmlArrayItem(Type = typeof(IdentityQueryAction), ElementName = "ObjectIdentityQuery")]
        public List<BaseAction> Actions { get; set; }
        [XmlArrayItem(Type = typeof(Identity)),
            XmlArrayItem(Type = typeof(Property)),
            XmlArrayItem(Type = typeof(StaticProperty)),
            XmlArrayItem(Type = typeof(ObjectPathMethod), ElementName = "Method"),
            XmlArrayItem(Type = typeof(StaticMethod), ElementName = "StaticMethod")]
        public List<Identity> ObjectPaths { get; set; }

    }
}
