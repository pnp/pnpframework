using Microsoft.SharePoint.Client;
using PnP.Framework.Provisioning.Connectors;

namespace PnP.Framework.Provisioning.Providers.Xml
{
    public class XMLSharePointTemplateProvider : XMLTemplateProvider
    {

        public XMLSharePointTemplateProvider() : base()
        {
        }

        public XMLSharePointTemplateProvider(ClientRuntimeContext cc, string connectionString, string container) :
            base(new SharePointConnector(cc, connectionString, container))
        {
        }
    }
}
