using Microsoft.SharePoint.Client;
using PnP.Framework.Attributes;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PnP.Framework.Provisioning.ObjectHandlers.TokenDefinitions
{
    [TokenDefinitionDescription(
       Token = "{authenticationrealm}",
       Description = "Returns the authentication ID of the current tenant/farm",
       Example = "{authenticationrealm}",
       Returns = "55898e77-a7bf-4799-8034-506db5521b98")]
    [TokenDefinitionDescription(
       Token = "{realm}",
       Description = "Returns the authentication ID of the current tenant/farm",
       Example = "{realm}",
       Returns = "55898e77-a7bf-4799-8034-506db5521b98")]
    internal class AuthenticationRealmToken : TokenDefinition
    {
        public AuthenticationRealmToken(Web web)
            : base(web, "{authenticationrealm}", "{realm}")
        {
        }
        public override string GetReplaceValue()
        {
            return Web.GetAuthenticationRealm().ToString();
        }
    }
}
