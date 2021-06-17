using System;
using System.Collections.Generic;
using System.Text;
using Microsoft.SharePoint.Client;
using PnP.Framework.Attributes;

namespace PnP.Framework.Provisioning.ObjectHandlers.TokenDefinitions
{
    [TokenDefinitionDescription(
        Token = "{sequencesitetoken:[provisioningid]:[siteTokenName]}",
        Description = "Returns the value of the named token in the specified site provisioning sequence",
        Example = "{sequencesitetoken:MYID:listid:My List}",
        Returns = "c7d9f9aa-4696-4c27-8a22-7d8eb7e70fda")]
    internal class SequenceSiteTokenToken : TokenDefinition
    {
        private string _value;
        public SequenceSiteTokenToken(Web web, string provisioningId, string siteTokenName, string siteTokenValue)
            : base(web, $"{{sequencesitetoken:{provisioningId}:{siteTokenName.Trim('{', '}')}}}")
        {
            _value = siteTokenValue;
        }

        public override string GetReplaceValue()
        {
            return _value;
        }
    }
}
