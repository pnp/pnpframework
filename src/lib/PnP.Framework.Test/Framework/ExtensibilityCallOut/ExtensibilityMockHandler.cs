using Microsoft.SharePoint.Client;
using PnP.Framework.Diagnostics;
using PnP.Framework.Provisioning.Extensibility;
using PnP.Framework.Provisioning.Model;
using PnP.Framework.Provisioning.ObjectHandlers;
using PnP.Framework.Provisioning.ObjectHandlers.TokenDefinitions;
using System;
using System.Collections.Generic;

namespace PnP.Framework.Tests.Framework.ExtensibilityCallOut
{
    public class ExtensibilityMockHandler : IProvisioningExtensibilityHandler
    {
        public ProvisioningTemplate Extract(ClientContext ctx, ProvisioningTemplate template, ProvisioningTemplateCreationInformation creationInformation, PnPMonitoredScope scope, string configurationData)
        {
            template.Lists.Add(new ListInstance() { Title = "Test List" });

            return template;
        }

        public IEnumerable<TokenDefinition> GetTokens(ClientContext ctx, ProvisioningTemplate template, string configurationData)
        {
            throw new NotImplementedException();
        }

        public void Provision(ClientContext ctx, ProvisioningTemplate template, ProvisioningTemplateApplyingInformation applyingInformation, TokenParser tokenParser, PnPMonitoredScope scope, string configurationData)
        {
            bool _urlCheck = ctx.Url.Equals(ExtensibilityTestConstants.MOCK_URL, StringComparison.OrdinalIgnoreCase);
            if (!_urlCheck) throw new Exception("CTXURLNOTTHESAME");

            bool _templateCheck = template.Id.Equals(ExtensibilityTestConstants.PROVISIONINGTEMPLATE_ID, StringComparison.OrdinalIgnoreCase);
            if (!_templateCheck) throw new Exception("TEMPLATEIDNOTTHESAME");

            bool _configDataCheck = configurationData.Equals(ExtensibilityTestConstants.PROVIDER_MOCK_DATA, StringComparison.OrdinalIgnoreCase);
            if (!_configDataCheck) throw new Exception("CONFIGDATANOTTHESAME");
        }
    }
}
