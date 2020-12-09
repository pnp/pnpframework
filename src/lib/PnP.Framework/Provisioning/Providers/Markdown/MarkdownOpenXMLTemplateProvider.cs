using PnP.Framework.Provisioning.Connectors;
using System;
using System.Security.Cryptography.X509Certificates;

namespace PnP.Framework.Provisioning.Providers.Markdown
{
    public class MarkdownOpenXMLTemplateProvider : MarkdownTemplateProvider
    {
        public MarkdownOpenXMLTemplateProvider(string packageFileName,
            FileConnectorBase persistenceConnector,
            String author = null,
            X509Certificate2 signingCertificate = null) :
            base(new OpenXMLConnector(packageFileName, persistenceConnector,
                author, signingCertificate))
        {
        }

        public MarkdownOpenXMLTemplateProvider(OpenXMLConnector openXMLConnector) :
            base(openXMLConnector)
        {
        }
    }
}
