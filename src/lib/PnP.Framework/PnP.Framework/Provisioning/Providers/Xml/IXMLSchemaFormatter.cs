using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;

namespace PnP.Framework.Provisioning.Providers.Xml
{
    /// <summary>
    /// Interface for template formatters that read and write XML documents
    /// </summary>
    public interface IXMLSchemaFormatter
    {
        /// <summary>
        /// The URI of the target XML Namespace
        /// </summary>
        String NamespaceUri { get; }

        /// <summary>
        /// The default namespace prefix of the target XML Namespace
        /// </summary>
        String NamespacePrefix { get; }
    }
}