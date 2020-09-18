using PnP.Framework.Provisioning.Connectors;

namespace PnP.Framework.Provisioning.Providers.Xml
{
    public class XMLFileSystemTemplateProvider : XMLTemplateProvider
    {

        public XMLFileSystemTemplateProvider(): base()
        {
        }

        public XMLFileSystemTemplateProvider(string connectionString, string container) :
            base(new FileSystemConnector(connectionString, container))
        {
        }
    }
}
