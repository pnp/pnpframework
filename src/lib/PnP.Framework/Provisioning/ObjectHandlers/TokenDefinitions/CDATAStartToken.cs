using Microsoft.SharePoint.Client;

namespace PnP.Framework.Provisioning.ObjectHandlers.TokenDefinitions
{
    internal class CDATAStartToken : TokenDefinition
    {
        public CDATAStartToken(Web web)
            : base(web, "{cdatastart}")
        {
        }

        public override string GetReplaceValue()
        {
            return "<![CDATA[";
        }
    }
}