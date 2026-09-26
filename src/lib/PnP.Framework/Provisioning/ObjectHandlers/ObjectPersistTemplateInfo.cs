using Microsoft.SharePoint.Client;
using Newtonsoft.Json;
using PnP.Framework.Diagnostics;
using PnP.Framework.Provisioning.Model;
using System;

namespace PnP.Framework.Provisioning.ObjectHandlers
{
    internal class ObjectPersistTemplateInfo : ObjectHandlerBase
    {
        public override string Name
        {
            get { return "Persist Template Info"; }
        }

        public override string InternalName => "PersistTemplateInfo";

        public ObjectPersistTemplateInfo()
        {
            this.ReportProgress = false;
        }

        public override TokenParser ProvisionObjects(Web web, ProvisioningTemplate template, TokenParser parser, ProvisioningTemplateApplyingInformation applyingInformation)
        {
            using (var scope = new PnPMonitoredScope(this.Name))
            {
                // Check if property bag writes are blocked (NoScript without tenant override)
                if (applyingInformation.PropertyBagWriteAllowed == false)
                {
                    return parser;
                }

                web.SetPropertyBagValue("_PnP_ProvisioningTemplateId", template.Id != null ? template.Id : "");
                web.AddIndexedPropertyBagKey("_PnP_ProvisioningTemplateId");

                ProvisioningTemplateInfo info = new ProvisioningTemplateInfo
                {
                    TemplateId = template.Id != null ? template.Id : "",
                    TemplateVersion = template.Version,
                    TemplateSitePolicy = template.SitePolicy,
                    Result = true,
                    ProvisioningTime = DateTime.Now
                };

                string jsonInfo = JsonConvert.SerializeObject(info);

                web.SetPropertyBagValue("_PnP_ProvisioningTemplateInfo", jsonInfo);
            }
            return parser;
        }

        public override Model.ProvisioningTemplate ExtractObjects(Web web, ProvisioningTemplate template, ProvisioningTemplateCreationInformation creationInfo)
        {
            return template;
        }

        public override bool WillProvision(Web web, ProvisioningTemplate template, ProvisioningTemplateApplyingInformation applyingInformation)
        {
            if (!_willProvision.HasValue)
            {
                _willProvision = applyingInformation.PropertyBagWriteAllowed != false;
            }
            return _willProvision.Value;
        }

        public override bool WillExtract(Web web, ProvisioningTemplate template, ProvisioningTemplateCreationInformation creationInfo)
        {
            if (!_willExtract.HasValue)
            {
                _willExtract = false;
            }
            return _willExtract.Value;
        }
    }
}
