using Microsoft.SharePoint.Client;
using PnP.Framework.Attributes;

namespace PnP.Framework.Provisioning.ObjectHandlers.TokenDefinitions
{
    [TokenDefinitionDescription(
    Token = "{currentuserloginname}",
    Description = "Returns the login name of the current user e.g. the user using the engine.",
    Example = "{currentuserloginname}",
    Returns = "i:0#.f|membership|user@domain.com")]
    internal class CurrentUserLoginNameToken : TokenDefinition
    {
        public CurrentUserLoginNameToken(Web web)
            : base(web, "{currentuserloginname}")
        {

        }

        public override string GetReplaceValue()
        {
            if (CacheValue == null)
            {
                var currentUser = TokenContext.Web.EnsureProperty(w => w.CurrentUser);
                CacheValue = currentUser.LoginName;
            }
            return CacheValue;
        }
    }
}