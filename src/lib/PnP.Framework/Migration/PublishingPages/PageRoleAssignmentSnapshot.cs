using System.Collections.Generic;

namespace PnP.Framework.Migration.PublishingPages
{
    public sealed class PageRoleAssignmentSnapshot
    {
        public string PrincipalLoginName { get; set; }

        public string PrincipalTitle { get; set; }

        public IList<string> RoleDefinitionNames { get; set; } = new List<string>();
    }
}
