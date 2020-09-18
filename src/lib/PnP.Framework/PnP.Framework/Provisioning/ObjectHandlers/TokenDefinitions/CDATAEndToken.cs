using Microsoft.SharePoint.Client;
using System;

namespace PnP.Framework.Provisioning.ObjectHandlers.TokenDefinitions
{
    internal class CDATAEndToken : TokenDefinition
    {
        public CDATAEndToken(Web web)
            : base(web, "{cdataend}")
        {
        }

        public override string GetReplaceValue()
        {
            return "]]>";
        }
    }
}