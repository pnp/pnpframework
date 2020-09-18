#if !NETSTANDARD2_0
using PnP.Framework.Provisioning.Connectors;

namespace PnP.Framework.Provisioning.Providers.Xml
{
    public class XMLAzureStorageTemplateProvider : XMLTemplateProvider
    {
        /// <summary>
        /// Default Constructor
        /// </summary>
        public XMLAzureStorageTemplateProvider() : base()
        {
        }

        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="connectionString"></param>
        /// <param name="container"></param>
        public XMLAzureStorageTemplateProvider(string connectionString, string container) :
            base(new AzureStorageConnector(connectionString, container))
        {
        }
    }
}
#endif