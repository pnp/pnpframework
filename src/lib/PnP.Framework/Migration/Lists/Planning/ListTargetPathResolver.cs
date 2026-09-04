using PnP.Framework.Migration.Topology;
using System;
using System.Collections.Generic;
using System.Linq;

namespace PnP.Framework.Migration.Lists.Planning
{
    internal sealed class ListTargetInventoryItem
    {
        public Guid ListId { get; set; }

        public string RootFolderServerRelativeUrl { get; set; }

        public string Title { get; set; }

        public int BaseTemplate { get; set; }

        public string OriginalIdentifier { get; set; }

        public string PlanDigest { get; set; }
    }

    internal sealed class ListTargetPathResolution
    {
        public string PreferredTargetRootFolderServerRelativeUrl { get; set; }

        public string PreferredTargetTitle { get; set; }

        public string TargetRootFolderServerRelativeUrl { get; set; }

        public string TargetTitle { get; set; }

        public bool CollisionResolved { get; set; }

        public string Reason { get; set; }

        public ListTargetInventoryItem ExistingOwnedTarget { get; set; }
    }

    /// <summary>
    /// Resolves planning-time List collisions without changing any ancestor path.
    /// A previously migration-owned List wins; otherwise only the colliding List
    /// URL leaf (and, when necessary, its display title) receives a stable suffix.
    /// </summary>
    internal static class ListTargetPathResolver
    {
        public static ListTargetPathResolution Resolve(
            ListMaterializationPlan plan,
            int sourceBaseTemplate,
            IEnumerable<ListTargetInventoryItem> targetInventory)
        {
            if (plan == null)
            {
                throw new ArgumentNullException(nameof(plan));
            }

            var inventory = (targetInventory ?? Enumerable.Empty<ListTargetInventoryItem>()).ToArray();
            var preferredPath = string.IsNullOrWhiteSpace(plan.PreferredTargetRootFolderServerRelativeUrl)
                ? plan.TargetRootFolderServerRelativeUrl
                : plan.PreferredTargetRootFolderServerRelativeUrl;
            var preferredTitle = string.IsNullOrWhiteSpace(plan.PreferredTargetTitle)
                ? plan.TargetTitle
                : plan.PreferredTargetTitle;

            var owned = inventory
                .Where(value => value.BaseTemplate == sourceBaseTemplate)
                .Where(value => string.Equals(value.OriginalIdentifier, plan.OriginalIdentifier, StringComparison.Ordinal))
                .Where(value => string.Equals(
                    value.PlanDigest,
                    ComputeDigestForTarget(plan, value.RootFolderServerRelativeUrl, value.Title),
                    StringComparison.OrdinalIgnoreCase))
                .OrderByDescending(value => PathEquals(value.RootFolderServerRelativeUrl, preferredPath))
                .ThenBy(value => value.RootFolderServerRelativeUrl, StringComparer.OrdinalIgnoreCase)
                .ThenBy(value => value.ListId)
                .FirstOrDefault();
            if (owned != null)
            {
                return new ListTargetPathResolution
                {
                    PreferredTargetRootFolderServerRelativeUrl = preferredPath,
                    PreferredTargetTitle = preferredTitle,
                    TargetRootFolderServerRelativeUrl = owned.RootFolderServerRelativeUrl,
                    TargetTitle = owned.Title,
                    CollisionResolved = !PathEquals(owned.RootFolderServerRelativeUrl, preferredPath)
                        || !string.Equals(owned.Title, preferredTitle, StringComparison.Ordinal),
                    Reason = "Reuse the deterministic migration-owned List carrying the exact source identity and semantic digest.",
                    ExistingOwnedTarget = owned
                };
            }

            var exactPath = inventory.FirstOrDefault(value => PathEquals(value.RootFolderServerRelativeUrl, preferredPath));
            var sameTitle = inventory.Where(value => string.Equals(value.Title, preferredTitle, StringComparison.OrdinalIgnoreCase)).ToArray();
            if (exactPath == null && sameTitle.Length == 0)
            {
                return new ListTargetPathResolution
                {
                    PreferredTargetRootFolderServerRelativeUrl = preferredPath,
                    PreferredTargetTitle = preferredTitle,
                    TargetRootFolderServerRelativeUrl = preferredPath,
                    TargetTitle = preferredTitle
                };
            }

            var occupiedPaths = inventory.Select(value => value.RootFolderServerRelativeUrl).ToList();
            if (exactPath == null)
            {
                // A same-title collision is still a user-visible List collision. Force
                // allocation at this List leaf while keeping every ancestor unchanged.
                occupiedPaths.Add(preferredPath);
            }
            var targetPath = TopologyTargetPathAllocator.AllocateServerRelativePath(
                preferredPath,
                plan.OriginalIdentifier,
                occupiedPaths);
            var targetTitle = sameTitle.Length == 0
                ? preferredTitle
                : TopologyTargetPathAllocator.AllocateSegment(
                    preferredTitle,
                    plan.OriginalIdentifier,
                    inventory.Select(value => value.Title),
                    maximumLength: 255);
            var reasons = new List<string>();
            if (exactPath != null)
            {
                reasons.Add("the preferred List path is occupied by a foreign or incompatible List");
            }
            if (sameTitle.Length > 0)
            {
                reasons.Add("the preferred List title is already in use");
            }
            return new ListTargetPathResolution
            {
                PreferredTargetRootFolderServerRelativeUrl = preferredPath,
                PreferredTargetTitle = preferredTitle,
                TargetRootFolderServerRelativeUrl = targetPath,
                TargetTitle = targetTitle,
                CollisionResolved = true,
                Reason = "Allocated a stable suffix only at the List node because " + string.Join(" and ", reasons) + "."
            };
        }

        private static string ComputeDigestForTarget(ListMaterializationPlan plan, string targetPath, string targetTitle)
        {
            var path = plan.TargetRootFolderServerRelativeUrl;
            var title = plan.TargetTitle;
            try
            {
                plan.TargetRootFolderServerRelativeUrl = targetPath;
                plan.TargetTitle = targetTitle;
                return ListMigrationPlanFactory.ComputePlanDigest(plan);
            }
            finally
            {
                plan.TargetRootFolderServerRelativeUrl = path;
                plan.TargetTitle = title;
            }
        }

        private static bool PathEquals(string left, string right)
        {
            return string.Equals(Normalize(left), Normalize(right), StringComparison.OrdinalIgnoreCase);
        }

        private static string Normalize(string value)
        {
            return Uri.UnescapeDataString(value ?? string.Empty).TrimEnd('/');
        }
    }
}
