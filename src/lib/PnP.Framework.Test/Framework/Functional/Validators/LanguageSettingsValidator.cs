using Microsoft.SharePoint.Client;
using PnP.Framework.Provisioning.Model;
using PnP.Framework.Provisioning.ObjectHandlers;
using PnP.Framework.Provisioning.Providers.Xml;
using System.Xml.Linq;

namespace PnP.Framework.Tests.Framework.Functional.Validators
{

    public class LanguageSettingsValidator : ValidatorBase
    {
        #region construction        
        public LanguageSettingsValidator() : base()
        {
            // optionally override schema version
            SchemaVersion = XMLConstants.PROVISIONING_SCHEMA_NAMESPACE_2015_12;
            XPathQuery = "/pnp:Templates/pnp:ProvisioningTemplate/pnp:SupportedUILanguages";
        }

        public LanguageSettingsValidator(ClientContext cc) : this()
        {
            this.cc = cc;
        }

        #endregion

        #region Validation logic
        public bool Validate(SupportedUILanguageCollection sourceLanguageSettings, SupportedUILanguageCollection targetLanguageSettings, TokenParser tokenParser)
        {
            ProvisioningTemplate sourcePt = new ProvisioningTemplate
            {
                SupportedUILanguages = sourceLanguageSettings
            };
            var sourceXml = ExtractElementXml(sourcePt);

            ProvisioningTemplate targetPt = new ProvisioningTemplate
            {
                SupportedUILanguages = targetLanguageSettings
            };
            var targetXml = ExtractElementXml(targetPt);

            return ValidateObjectXML(sourceXml, targetXml, null);
        }

        internal override void OverrideXmlData(XElement sourceObject, XElement targetObject)
        {
            XNamespace ns = SchemaVersion;

        }
        #endregion

        #region Helper methods
        #endregion
    }
}
