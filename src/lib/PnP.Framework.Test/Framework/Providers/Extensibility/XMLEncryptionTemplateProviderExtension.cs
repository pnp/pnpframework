using PnP.Framework.Provisioning.Model;
using PnP.Framework.Provisioning.Providers;
using PnP.Framework.Provisioning.Providers.Xml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Security.Cryptography.X509Certificates;

namespace PnP.Framework.Test.Framework.Providers.Extensibility
{
    public class XMLEncryptionTemplateProviderExtension : ITemplateProviderExtension
    {
        public bool SupportsGetTemplatePostProcessing
        {
            get
            {
                return (false);
            }
        }

        public bool SupportsGetTemplatePreProcessing
        {
            get
            {
                return (true);
            }
        }

        public bool SupportsSaveTemplatePostProcessing
        {
            get
            {
                return (true);
            }
        }

        public bool SupportsSaveTemplatePreProcessing
        {
            get
            {
                return (false);
            }
        }

        private X509Certificate2 _certificate;

        public void Initialize(object settings)
        {
            _certificate = settings as X509Certificate2;
        }

        public ProvisioningTemplate PostProcessGetTemplate(ProvisioningTemplate template)
        {
            throw new NotImplementedException();
        }

        public Stream PostProcessSaveTemplate(Stream stream)
        {
            MemoryStream result = new MemoryStream();

            var namespaces = new Dictionary<string, string>
            {
                { "pnp", XMLConstants.PROVISIONING_SCHEMA_NAMESPACE_2022_09 }
            };

            SecureXml.EncryptXmlDocument(stream, result, "/pnp:Provisioning", namespaces, this._certificate);
            result.Position = 0;

            return (result);
        }

        public Stream PreProcessGetTemplate(Stream stream)
        {
            MemoryStream result = new MemoryStream();

            SecureXml.DecryptXmlDocument(stream, result, this._certificate);
            result.Position = 0;

            return (result);
        }

        public ProvisioningTemplate PreProcessSaveTemplate(ProvisioningTemplate template)
        {
            throw new NotImplementedException();
        }
    }
}
