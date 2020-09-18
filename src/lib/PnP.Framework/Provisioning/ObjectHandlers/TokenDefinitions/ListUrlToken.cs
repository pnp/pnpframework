using Microsoft.SharePoint.Client;
using PnP.Framework.Attributes;
using System.Text.RegularExpressions;

namespace PnP.Framework.Provisioning.ObjectHandlers.TokenDefinitions
{
    [TokenDefinitionDescription(
     Token = "{listurl:[name]}",
     Description = "Returns a site relative url of the list given its name",
     Example = "{listid:My List}",
     Returns = "Lists/MyList")]
    internal class ListUrlToken : TokenDefinition
    {
        private readonly string _listUrl = null;
        public ListUrlToken(Web web, string name, string url)
            : base(web, $"{{listurl:{Regex.Escape(name)}}}")
        {
            _listUrl = url;
        }

        public override string GetReplaceValue()
        {
            if (string.IsNullOrEmpty(CacheValue))
            {
                CacheValue = _listUrl;
            }
            return CacheValue;
        }
    }
}