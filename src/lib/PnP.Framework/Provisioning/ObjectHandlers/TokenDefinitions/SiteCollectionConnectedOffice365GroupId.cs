using Microsoft.SharePoint.Client;
using PnP.Framework.Attributes;
using System;

namespace PnP.Framework.Provisioning.ObjectHandlers.TokenDefinitions
{
    [TokenDefinitionDescription(
        Token = "{sitecollectionconnectedoffice365groupid}",
        Description = "Returns the ID of the Office 365 group connected to the current site",
        Example = "{sitecollectionconnectedoffice365groupid}",
        Returns = "767bc144-e605-4d8c-885a-3a980feb39c6")]
    internal class SiteCollectionConnectedOffice365GroupId : VolatileTokenDefinition
    {
        public SiteCollectionConnectedOffice365GroupId(Web web)
            : base(web, "{sitecollectionconnectedoffice365groupid}")
        {
        }

        public override string GetReplaceValue()
        {
            if (CacheValue == null)
            {
                TokenContext.Load(TokenContext.Site, s => s.GroupId);
                TokenContext.ExecuteQueryRetry();
                if (!TokenContext.Site.GroupId.Equals(Guid.Empty))
                {
                    CacheValue = TokenContext.Site.GroupId.ToString();
                }
                else
                {
                    CacheValue = "";
                }
            }
            return CacheValue;
        }
    }
}
