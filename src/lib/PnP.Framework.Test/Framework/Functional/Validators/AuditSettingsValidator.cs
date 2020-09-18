using Microsoft.SharePoint.Client;
using PnP.Framework.Provisioning.Model;
using PnP.Framework.Provisioning.ObjectHandlers;
using PnP.Framework.Provisioning.Providers.Xml;
using System;
using System.Xml.Linq;

namespace PnP.Framework.Test.Framework.Functional.Validators
{

    public class AuditSettingsValidator : ValidatorBase
    {
        private readonly bool isNoScriptSite = false;

        #region construction        
        public AuditSettingsValidator() : base()
        {
            // optionally override schema version
            SchemaVersion = XMLConstants.PROVISIONING_SCHEMA_NAMESPACE_2015_12;
            XPathQuery = "/pnp:Templates/pnp:ProvisioningTemplate/pnp:AuditSettings";
        }

        public AuditSettingsValidator(ClientContext cc) : this()
        {
            this.cc = cc;
            isNoScriptSite = cc.Web.IsNoScriptSite();
        }

        #endregion

        #region Validation logic
        public bool Validate(AuditSettings sourceAuditsettings, AuditSettings targetAuditSettings, TokenParser tokenParser)
        {
            ProvisioningTemplate sourcePt = new ProvisioningTemplate
            {
                AuditSettings = sourceAuditsettings
            };
            var sourceXml = ExtractElementXml(sourcePt);

            ProvisioningTemplate targetPt = new ProvisioningTemplate
            {
                AuditSettings = targetAuditSettings
            };
            var targetXml = ExtractElementXml(targetPt);

            return ValidateObjectXML(sourceXml, targetXml, null);
        }

        internal override void OverrideXmlData(XElement sourceObject, XElement targetObject)
        {
            XNamespace ns = SchemaVersion;

            if (isNoScriptSite)
            {
                DropAttribute(sourceObject, "AuditLogTrimmingRetention");
                DropAttribute(targetObject, "AuditLogTrimmingRetention");
            }
        }
        #endregion

        #region Helper methods
        private bool ValidateMasterPage(string source, string target)
        {
            if (!source.StartsWith("/_catalogs/MasterPage", StringComparison.InvariantCultureIgnoreCase))
            {
                int loc = source.IndexOf("/_catalogs");
                if (loc >= 0)
                {
                    source = source.Substring(loc);

                    if (!source.Equals(target, StringComparison.InvariantCultureIgnoreCase))
                    {
                        return false;
                    }
                }
            }

            return true;
        }
        #endregion
    }
}
