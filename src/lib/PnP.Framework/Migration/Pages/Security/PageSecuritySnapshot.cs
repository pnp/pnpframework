using System.Collections.Generic;

namespace PnP.Framework.Migration.Pages.Security
{
    public sealed class PageSecuritySnapshot
    {
        public bool HasUniqueRoleAssignments { get; set; }

        public IList<PageRoleAssignmentSnapshot> RoleAssignments { get; set; } = new List<PageRoleAssignmentSnapshot>();
    }
}
