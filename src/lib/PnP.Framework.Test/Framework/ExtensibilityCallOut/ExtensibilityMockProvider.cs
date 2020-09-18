using Microsoft.SharePoint.Client;
using PnP.Framework.Provisioning.Extensibility;
using PnP.Framework.Provisioning.Model;
using System;

namespace PnP.Framework.Tests.Framework.ExtensibilityCallOut
{
    public class ExtensibilityMockProvider : IProvisioningExtensibilityProvider
    {
        public void ProcessRequest(ClientContext ctx, ProvisioningTemplate template, string configurationData)
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
