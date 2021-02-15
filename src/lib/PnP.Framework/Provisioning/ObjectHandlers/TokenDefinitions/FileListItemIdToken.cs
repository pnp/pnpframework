using Microsoft.SharePoint.Client;
using PnP.Framework.Attributes;
using System.Text.RegularExpressions;

namespace PnP.Framework.Provisioning.ObjectHandlers.TokenDefinitions
{
    [TokenDefinitionDescription(
     Token = "{filelistitemid:[siteRelativePath]}",
     Description = "Returns the listitem id of a file which is being provisioned by the current template.",
     Example = "{filelistitemid:/library/folder/file.docx}",
     Returns = "54")]
    internal class FileListItemIdToken : TokenDefinition
    {
        private readonly string _value = null;
        public FileListItemIdToken(Web web, string siteRelativePath, int id)
            : base(web, $"{{filelistitemid:{Regex.Escape(siteRelativePath)}}}")
        {
            _value = id.ToString();
        }

        public override string GetReplaceValue()
        {
            if (string.IsNullOrEmpty(CacheValue))
            {
                CacheValue = _value;
            }
            return CacheValue;
        }
    }
}

