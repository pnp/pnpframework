using Microsoft.SharePoint.Client;
using PnP.Framework.Migration.Diagnostics;
using PnP.Framework.Migration.Lists.Capture;
using System;
using System.Collections.Generic;
using System.Linq;

namespace PnP.Framework.Migration.Lists.Planning
{
    internal static class ListTargetInspector
    {
        public const string OriginalIdentifierPropertyName = "pnp_reserved_list_original_identifier";
        public const string PlanDigestPropertyName = "pnp_reserved_list_migration_digest";

        public static ListTargetProbe Inspect(ClientContext anchorContext, ListDependencySnapshot source, ListMaterializationPlan plan)
        {
            return Inspect(anchorContext, source, plan, false);
        }

        public static ListTargetProbe InspectForPlanning(ClientContext anchorContext, ListDependencySnapshot source, ListMaterializationPlan plan)
        {
            return Inspect(anchorContext, source, plan, true);
        }

        private static ListTargetProbe Inspect(
            ClientContext anchorContext,
            ListDependencySnapshot source,
            ListMaterializationPlan plan,
            bool resolvePlanningCollision)
        {
            if (anchorContext == null)
            {
                throw new ArgumentNullException(nameof(anchorContext));
            }
            var probe = new ListTargetProbe
            {
                PreferredTargetRootFolderServerRelativeUrl = plan.PreferredTargetRootFolderServerRelativeUrl ?? plan.TargetRootFolderServerRelativeUrl,
                PreferredTargetTitle = plan.PreferredTargetTitle ?? plan.TargetTitle,
                TargetWebUrl = plan.TargetWebUrl,
                TargetRootFolderServerRelativeUrl = plan.TargetRootFolderServerRelativeUrl,
                TargetTitle = plan.TargetTitle,
                Disposition = ListMaterializationDisposition.Block
            };
            try
            {
                if (CanUseLoadedWeb(anchorContext.Web, plan.TargetWebUrl))
                {
                    return PopulateProbe(anchorContext.Web, source, plan, probe, resolvePlanningCollision);
                }

                using (var context = anchorContext.Clone(plan.TargetWebUrl))
                {
                    var web = context.Web;
                    context.Load(web, value => value.Id, value => value.Url, value => value.ServerRelativeUrl, value => value.EffectiveBasePermissions);
                    context.Load(web.Lists, values => values.Include(
                        value => value.Id,
                        value => value.Title,
                        value => value.BaseTemplate,
                        value => value.RootFolder.ServerRelativeUrl,
                        value => value.RootFolder.Properties));
                    context.ExecuteQueryRetry();
                    return PopulateProbe(web, source, plan, probe, resolvePlanningCollision);
                }
            }
            catch (Exception exception) when (exception is ServerException || exception is ClientRequestException)
            {
                probe.Issues.Add(Issue("TargetWebUnavailable", plan, "The mapped target Web could not be inspected: " + exception.Message));
                return probe;
            }
        }

        private static bool CanUseLoadedWeb(Web web, string targetWebUrl)
        {
            return web != null
                && web.IsPropertyAvailable("Id")
                && web.IsPropertyAvailable("Url")
                && web.IsPropertyAvailable("EffectiveBasePermissions")
                && web.Lists.AreItemsAvailable
                && string.Equals(
                    new Uri(web.Url).AbsoluteUri.TrimEnd('/'),
                    new Uri(targetWebUrl).AbsoluteUri.TrimEnd('/'),
                    StringComparison.OrdinalIgnoreCase);
        }

        private static ListTargetProbe PopulateProbe(
            Web web,
            ListDependencySnapshot source,
            ListMaterializationPlan plan,
            ListTargetProbe probe,
            bool resolvePlanningCollision)
        {
            probe.TargetWebExists = true;
            probe.TargetWebId = web.Id;
            probe.CanManageLists = web.EffectiveBasePermissions.Has(PermissionKind.ManageLists);
            if (resolvePlanningCollision)
            {
                var resolution = ListTargetPathResolver.Resolve(
                    plan,
                    source.BaseTemplate,
                    web.Lists.AsEnumerable().Select(value => new ListTargetInventoryItem
                    {
                        ListId = value.Id,
                        RootFolderServerRelativeUrl = value.RootFolder.ServerRelativeUrl,
                        Title = value.Title,
                        BaseTemplate = value.BaseTemplate,
                        OriginalIdentifier = Property(value.RootFolder.Properties, OriginalIdentifierPropertyName),
                        PlanDigest = Property(value.RootFolder.Properties, PlanDigestPropertyName)
                    }));
                probe.PreferredTargetRootFolderServerRelativeUrl = resolution.PreferredTargetRootFolderServerRelativeUrl;
                probe.PreferredTargetTitle = resolution.PreferredTargetTitle;
                probe.TargetRootFolderServerRelativeUrl = resolution.TargetRootFolderServerRelativeUrl;
                probe.TargetTitle = resolution.TargetTitle;
                probe.CollisionResolved = resolution.CollisionResolved;
                probe.CollisionResolutionReason = resolution.Reason;
                if (resolution.ExistingOwnedTarget != null)
                {
                    probe.ListExists = true;
                    probe.TargetListId = resolution.ExistingOwnedTarget.ListId;
                    probe.ExistingTitle = resolution.ExistingOwnedTarget.Title;
                    probe.ExistingBaseTemplate = resolution.ExistingOwnedTarget.BaseTemplate;
                    probe.ExistingOriginalIdentifier = resolution.ExistingOwnedTarget.OriginalIdentifier;
                    probe.ExistingPlanDigest = resolution.ExistingOwnedTarget.PlanDigest;
                    probe.Disposition = ListMaterializationDisposition.ReuseOwned;
                    return probe;
                }
                if (!probe.CanManageLists)
                {
                    probe.Issues.Add(Issue("TargetListWriteUnavailable", plan, "The mapped target Web does not grant ManageLists."));
                    probe.Disposition = ListMaterializationDisposition.Block;
                    return probe;
                }
                probe.Disposition = ListMaterializationDisposition.CreateOwned;
                return probe;
            }

            var exact = web.Lists.AsEnumerable().FirstOrDefault(value => string.Equals(
                Normalize(value.RootFolder.ServerRelativeUrl),
                Normalize(plan.TargetRootFolderServerRelativeUrl),
                StringComparison.OrdinalIgnoreCase));
            probe.SameTitleDifferentPaths = web.Lists.AsEnumerable()
                .Where(value => string.Equals(value.Title, plan.TargetTitle, StringComparison.OrdinalIgnoreCase))
                .Where(value => exact == null || value.Id != exact.Id)
                .Select(value => value.RootFolder.ServerRelativeUrl)
                .OrderBy(value => value, StringComparer.OrdinalIgnoreCase).ToList();
            if (exact == null)
            {
                if (!probe.CanManageLists)
                {
                    probe.Issues.Add(Issue("TargetListWriteUnavailable", plan, "The mapped target Web does not grant ManageLists."));
                }
                if (probe.SameTitleDifferentPaths.Count > 0)
                {
                    probe.Issues.Add(Issue("TargetListTitleCollision", plan,
                        "The target already contains the same List title at different path(s): " + string.Join(", ", probe.SameTitleDifferentPaths) + "."));
                }
                probe.Disposition = probe.Issues.Count == 0 ? ListMaterializationDisposition.CreateOwned : ListMaterializationDisposition.Block;
                return probe;
            }

            probe.ListExists = true;
            probe.TargetListId = exact.Id;
            probe.ExistingTitle = exact.Title;
            probe.ExistingBaseTemplate = exact.BaseTemplate;
            probe.ExistingOriginalIdentifier = Property(exact.RootFolder.Properties, OriginalIdentifierPropertyName);
            probe.ExistingPlanDigest = Property(exact.RootFolder.Properties, PlanDigestPropertyName);
            if (exact.BaseTemplate != source.BaseTemplate || !string.Equals(exact.Title, plan.TargetTitle, StringComparison.Ordinal))
            {
                probe.Issues.Add(Issue("TargetListCollision", plan, "A List exists at the target path with different title or base template."));
            }
            if (!string.Equals(probe.ExistingOriginalIdentifier, plan.OriginalIdentifier, StringComparison.Ordinal)
                || !string.Equals(probe.ExistingPlanDigest, plan.PlanDigest, StringComparison.OrdinalIgnoreCase))
            {
                probe.Issues.Add(Issue("TargetListOwnershipCollision", plan, "The existing List is not claimed by this exact source identity and semantic plan digest."));
            }
            probe.Disposition = probe.Issues.Count == 0 ? ListMaterializationDisposition.ReuseOwned : ListMaterializationDisposition.Block;
            return probe;
        }

        public static ListTargetProbe DeferUntilTopologyMaterialization(ListMaterializationPlan plan)
        {
            if (plan == null)
            {
                throw new ArgumentNullException(nameof(plan));
            }
            return new ListTargetProbe
            {
                PreferredTargetRootFolderServerRelativeUrl = plan.PreferredTargetRootFolderServerRelativeUrl ?? plan.TargetRootFolderServerRelativeUrl,
                PreferredTargetTitle = plan.PreferredTargetTitle ?? plan.TargetTitle,
                TargetWebUrl = plan.TargetWebUrl,
                TargetRootFolderServerRelativeUrl = plan.TargetRootFolderServerRelativeUrl,
                TargetTitle = plan.TargetTitle,
                TargetWebExists = false,
                DeferredUntilTopologyMaterialization = true,
                CanManageLists = true,
                Disposition = ListMaterializationDisposition.CreateOwned
            };
        }

        private static string Normalize(string value)
        {
            return Uri.UnescapeDataString(value ?? string.Empty).TrimEnd('/');
        }

        private static string Property(PropertyValues values, string name)
        {
            object value;
            return values != null && values.FieldValues.TryGetValue(name, out value) ? Convert.ToString(value) : null;
        }

        private static MigrationIssue Issue(string code, ListMaterializationPlan plan, string message)
        {
            return new MigrationIssue
            {
                Code = code,
                Severity = MigrationIssueSeverity.Blocker,
                Subject = "target-list:" + plan.TargetRootFolderServerRelativeUrl,
                Ingredient = "ListDependency.TargetProbe",
                Message = message,
                SourceIdentity = plan.OriginalIdentifier,
                TargetIdentity = plan.TargetRootFolderServerRelativeUrl
            };
        }
    }
}
