using Microsoft.SharePoint.Client;
using PnP.Framework.Attributes;

namespace PnP.Framework.Provisioning.ObjectHandlers.TokenDefinitions
{
    [TokenDefinitionDescription(
      Token = "{currentuserfullname}",
      Description = "Returns the full name of the current user e.g. the user using the engine.",
      Example = "{currentuserfullname}",
      Returns = "John Doe")]
    internal class CurrentUserFullNameToken : TokenDefinition
    {
        public CurrentUserFullNameToken(Web web)
            : base(web, "{currentuserfullname}")
        {
        }

        public override string GetReplaceValue()
        {
            if (CacheValue == null)
            {
                var currentUser = TokenContext.Web.EnsureProperty(w => w.CurrentUser);
                CacheValue = currentUser.Title;
            }
            return CacheValue;
        }
    }
}