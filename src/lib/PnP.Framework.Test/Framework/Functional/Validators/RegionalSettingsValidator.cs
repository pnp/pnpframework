using Microsoft.SharePoint.Client;
using PnP.Framework.Provisioning.Model;
using PnP.Framework.Provisioning.ObjectHandlers;
using PnP.Framework.Provisioning.Providers.Xml;
using System.Xml.Linq;

namespace PnP.Framework.Tests.Framework.Functional.Validators
{

    public class RegionalSettingsValidator : ValidatorBase
    {
        #region construction        
        public RegionalSettingsValidator() : base()
        {
            // optionally override schema version
            SchemaVersion = XMLConstants.PROVISIONING_SCHEMA_NAMESPACE_2015_12;
            XPathQuery = "/pnp:Templates/pnp:ProvisioningTemplate/pnp:RegionalSettings";
        }

        public RegionalSettingsValidator(ClientContext cc) : this()
        {
            this.cc = cc;
        }

        #endregion

        #region Validation logic
        public bool Validate(PnP.Framework.Provisioning.Model.RegionalSettings sourceRegionalSettings, PnP.Framework.Provisioning.Model.RegionalSettings targetRegionalSettings, TokenParser tokenParser)
        {
            ProvisioningTemplate sourcePt = new ProvisioningTemplate
            {
                RegionalSettings = sourceRegionalSettings
            };
            var sourceXml = ExtractElementXml(sourcePt);

            ProvisioningTemplate targetPt = new ProvisioningTemplate
            {
                RegionalSettings = targetRegionalSettings
            };
            var targetXml = ExtractElementXml(targetPt);

            return ValidateObjectXML(sourceXml, targetXml, null, null, null);
        }

        internal override void OverrideXmlData(XElement sourceObject, XElement targetObject)
        {
            XNamespace ns = SchemaVersion;

            if (sourceObject.Attribute("CalendarType").Value.ToLower() == "none")
            {
                DropAttribute(sourceObject, "CalendarType");
                DropAttribute(targetObject, "CalendarType");
            }

        }

        #endregion
    }
}
