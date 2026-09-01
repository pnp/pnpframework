using Microsoft.SharePoint.Client;
using System;
using System.Collections.Generic;
using System.Linq;

namespace PnP.Framework.Migration.Pages.Security
{
    internal static class PageSecuritySnapshotReader
    {
        public static PageSecuritySnapshot Read(ClientContext context, ListItem item, ICollection<string> warnings)
        {
            var result = new PageSecuritySnapshot
            {
                HasUniqueRoleAssignments = item.HasUniqueRoleAssignments
            };
            if (!item.HasUniqueRoleAssignments)
            {
                return result;
            }

            try
            {
                var assignments = item.RoleAssignments;
                context.Load(assignments);
                context.ExecuteQueryRetry();
                foreach (var assignment in assignments)
                {
                    context.Load(assignment.Member, member => member.LoginName, member => member.Title);
                    context.Load(assignment.RoleDefinitionBindings, definitions => definitions.Include(definition => definition.Name));
                }

                context.ExecuteQueryRetry();
                foreach (var assignment in assignments)
                {
                    result.RoleAssignments.Add(new PageRoleAssignmentSnapshot
                    {
                        PrincipalLoginName = assignment.Member.LoginName,
                        PrincipalTitle = assignment.Member.Title,
                        RoleDefinitionNames = assignment.RoleDefinitionBindings
                            .Select(definition => definition.Name)
                            .OrderBy(name => name, StringComparer.OrdinalIgnoreCase)
                            .ToList()
                    });
                }
            }
            catch (Exception exception) when (IsAccessDenied(exception))
            {
                warnings.Add("The source page has unique permissions, but the current principal cannot enumerate its role assignments. Permission replay is not supported by this migration profile, so page capture continued without ACL details.");
            }

            return result;
        }

        private static bool IsAccessDenied(Exception exception)
        {
            for (var current = exception; current != null; current = current.InnerException)
            {
                if (current is UnauthorizedAccessException || current is ServerUnauthorizedAccessException)
                {
                    return true;
                }

                if (current is ServerException serverException && serverException.ServerErrorCode == -2147024891)
                {
                    return true;
                }
            }

            return false;
        }
    }
}
