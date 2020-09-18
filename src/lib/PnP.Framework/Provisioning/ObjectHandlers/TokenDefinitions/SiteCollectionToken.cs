using Microsoft.SharePoint.Client;
using PnP.Framework.Attributes;

namespace PnP.Framework.Provisioning.ObjectHandlers.TokenDefinitions
{
    [TokenDefinitionDescription(
        Token = "{sitecollection}",
        Description = "Returns the server relative url of the site collection",
        Example = "{sitecollection}",
        Returns = "/sites/mysitecollection")]
    internal class SiteCollectionToken : VolatileTokenDefinition
    {
        public SiteCollectionToken(Web web)
            : base(web, "{sitecollection}")
        {
        }

        public override string GetReplaceValue()
        {
            if (CacheValue == null)
            {
                var site = TokenContext.Site;
                TokenContext.Load(site, s => s.RootWeb.ServerRelativeUrl);
                TokenContext.ExecuteQueryRetry();
                CacheValue = site.RootWeb.ServerRelativeUrl.TrimEnd('/');
            }
            return CacheValue;
        }
    }
}