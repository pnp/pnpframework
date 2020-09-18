using Microsoft.SharePoint.Client;
using PnP.Framework.Attributes;
using System.Text.RegularExpressions;

namespace PnP.Framework.Provisioning.ObjectHandlers.TokenDefinitions
{
    [TokenDefinitionDescription(
      Token = "{contenttypeid:[contenttypename]}",
      Description = "Returns the ID of the specified content type",
      Example = "{contenttypeid:My Content Type}",
      Returns = "0x0102004F51EFDEA49C49668EF9C6744C8CF87D")]
    internal class ContentTypeIdToken : TokenDefinition
    {
        private readonly string _contentTypeId = null;
        public ContentTypeIdToken(Web web, string name, string contenttypeid)
            : base(web, $"{{contenttypeid:{Regex.Escape(name)}}}")
        {
            _contentTypeId = contenttypeid;
        }

        public override string GetReplaceValue()
        {
            if (string.IsNullOrEmpty(CacheValue))
            {
                CacheValue = _contentTypeId;
            }
            return CacheValue;
        }
    }
}