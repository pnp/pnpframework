using PnP.Framework.Provisioning.Connectors;
using System;
using System.Security.Cryptography.X509Certificates;

namespace PnP.Framework.Provisioning.Providers.Json
{
    public class JsonOpenXMLTemplateProvider : JsonTemplateProvider
    {
        public JsonOpenXMLTemplateProvider(string packageFileName,
            FileConnectorBase persistenceConnector,
            String author = null,
            X509Certificate2 signingCertificate = null) :
            base(new OpenXMLConnector(packageFileName, persistenceConnector,
                author, signingCertificate))
        {
        }

        public JsonOpenXMLTemplateProvider(OpenXMLConnector openXMLConnector) :
            base(openXMLConnector)
        {
        }
    }
}
