using Microsoft.SharePoint.Client;
using PnP.Framework.Attributes;
using System;
using System.Text.RegularExpressions;

namespace PnP.Framework.Provisioning.ObjectHandlers.TokenDefinitions
{
    [TokenDefinitionDescription(
      Token = "{sitescriptid:[scripttitle]}",
      Description = "Returns the id of the given site script",
      Example = "{sitescriptid:My Site Script}",
      Returns = "9188a794-cfcf-48b6-9ac5-df2048e8aa5d")]
    internal class SiteScriptIdToken : TokenDefinition
    {
        private Guid _scriptId;
        public SiteScriptIdToken(Web web, string scriptTitle, Guid scriptId)
            : base(web, $"{{sitescriptid:{Regex.Escape(scriptTitle)}}}")
        {
            _scriptId = scriptId;
        }

        public override string GetReplaceValue()
        {
            if (CacheValue == null)
            {
                CacheValue = _scriptId.ToString();
            }
            return CacheValue;
        }
    }
}