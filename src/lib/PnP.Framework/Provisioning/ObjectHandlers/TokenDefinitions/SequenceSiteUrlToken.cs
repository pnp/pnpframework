using Microsoft.SharePoint.Client;
using PnP.Framework.Attributes;

namespace PnP.Framework.Provisioning.ObjectHandlers.TokenDefinitions
{
    [TokenDefinitionDescription(
    Token = "{sequencesiteurl:[provisioningid]}",
    Description = "Returns a full url of the site given its provisioning ID from the sequence",
    Example = "{sequencesiteurl:MYID}",
    Returns = "https://contoso.sharepoint.com/sites/mynewsite")]
    internal class SequenceSiteUrlUrlToken : TokenDefinition
    {
        private readonly string _url = null;
        public SequenceSiteUrlUrlToken(Web web, string provisioningId, string url)
            : base(web, $"{{sequencesiteurl:{provisioningId}}}")
        {
            _url = url;
        }

        public override string GetReplaceValue()
        {
            if (string.IsNullOrEmpty(CacheValue))
            {
                CacheValue = _url;
            }
            return CacheValue;
        }
    }
}
