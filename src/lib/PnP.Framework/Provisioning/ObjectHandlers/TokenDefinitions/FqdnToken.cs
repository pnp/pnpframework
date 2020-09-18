using Microsoft.SharePoint.Client;
using PnP.Framework.Attributes;
using System;

namespace PnP.Framework.Provisioning.ObjectHandlers.TokenDefinitions
{
    [TokenDefinitionDescription(
     Token = "{fqdn}",
     Description = "Returns a full url of the current host",
     Example = "{fqdn}",
     Returns = "mycompany.sharepoint.com")]
    public class FqdnToken : TokenDefinition
    {
        public FqdnToken(Web web) : base(web, "{fqdn}")
        {
        }

        public override string GetReplaceValue()
        {
            if (CacheValue == null)
            {
                TokenContext.Web.EnsureProperty(w => w.Url);
                var uri = new Uri(TokenContext.Web.Url);
                CacheValue = $"{uri.DnsSafeHost.ToLower().Replace("-admin", "")}";
            }
            return CacheValue;
        }
    }
}
