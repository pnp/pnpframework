using Microsoft.VisualStudio.TestTools.UnitTesting;
using PnP.Framework.Provisioning.Model;
using PnP.Framework.Provisioning.ObjectHandlers;
using System.Xml.Linq;

namespace PnP.Framework.Tests.Framework.Functional.Validators
{
    [TestClass]
    public class SitePolicyValidator : ValidatorBase
    {
        public SitePolicyValidator() : base()
        {
            // optionally override schema version
            //SchemaVersion = XMLConstants.PROVISIONING_SCHEMA_NAMESPACE_2016_05;
        }

        public bool Validate(string source, string target, TokenParser parser)
        {
            XPathQuery = "/pnp:Templates/pnp:ProvisioningTemplate/pnp:SitePolicy";

            ProvisioningTemplate pt = new ProvisioningTemplate
            {
                SitePolicy = source
            };
            string sSchemaXml = ExtractElementXml(pt);

            ProvisioningTemplate ptTarget = new ProvisioningTemplate
            {
                SitePolicy = target
            };
            string tSchemaXml = ExtractElementXml(ptTarget);

            // Use XML validation logic to compare source and target
            if (!ValidateObjectXML(sSchemaXml, tSchemaXml, null)) { return false; }

            return true;
        }

        internal override void OverrideXmlData(XElement sourceObject, XElement targetObject)
        {
            XNamespace ns = SchemaVersion;
        }


    }
}
