using Microsoft.SharePoint.Client;
using PnP.Framework.Provisioning.Connectors;


namespace PnP.Framework.Provisioning.Providers.Json
{
    public class JsonSharePointTemplateProvider : JsonTemplateProvider
    {
        public JsonSharePointTemplateProvider() : base()
        {

        }

        public JsonSharePointTemplateProvider(ClientRuntimeContext cc, string connectionString, string container) :
            base(new SharePointConnector(cc, connectionString, container))
        {
        }
    }
}
