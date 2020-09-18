using Microsoft.SharePoint.Client;
using PnP.Framework.Attributes;

namespace PnP.Framework.Provisioning.ObjectHandlers.TokenDefinitions
{
    [TokenDefinitionDescription(
    Token = "{currentuserid}",
    Description = "Returns the ID of the current user e.g. the user using the engine.",
    Example = "{currentuserid}",
    Returns = "4")]
    internal class CurrentUserIdToken : VolatileTokenDefinition
    {
        public CurrentUserIdToken(Web web)
            : base(web, "{currentuserid}")
        {
        }

        public override string GetReplaceValue()
        {
            if (CacheValue == null)
            {
                var currentUser = TokenContext.Web.EnsureProperty(w => w.CurrentUser);
                CacheValue = currentUser.Id.ToString();
            }
            return CacheValue;
        }
    }
}