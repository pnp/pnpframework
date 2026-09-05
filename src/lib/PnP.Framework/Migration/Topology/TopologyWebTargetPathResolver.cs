using System;
using System.Collections.Generic;
using System.Linq;

namespace PnP.Framework.Migration.Topology
{
    internal sealed class TopologyWebTargetInventoryItem
    {
        public Guid WebId { get; set; }

        public string Url { get; set; }

        public string ServerRelativeUrl { get; set; }

        public string Title { get; set; }

        public string Description { get; set; }

        public string Template { get; set; }

        public int Configuration { get; set; }

        public string OriginalIdentifier { get; set; }

        public string MappingDigest { get; set; }
    }

    internal sealed class TopologyWebTargetPathResolution
    {
        public string PreferredTargetWebUrl { get; set; }

        public string PreferredTargetServerRelativeUrl { get; set; }

        public string TargetWebUrl { get; set; }

        public string TargetServerRelativeUrl { get; set; }

        public bool CollisionResolved { get; set; }

        public string Reason { get; set; }

        public TopologyWebTargetInventoryItem ExistingTarget { get; set; }

        public TopologyMaterializationDisposition ExistingDisposition { get; set; }
    }

    internal static class TopologyWebTargetPathResolver
    {
        public static TopologyWebTargetPathResolution Resolve(
            WebMappingPlan plan,
            IEnumerable<TopologyWebTargetInventoryItem> targetInventory)
        {
            if (plan == null)
            {
                throw new ArgumentNullException(nameof(plan));
            }

            var inventory = (targetInventory ?? Enumerable.Empty<TopologyWebTargetInventoryItem>()).ToArray();
            var preferredPath = string.IsNullOrWhiteSpace(plan.PreferredTargetServerRelativeUrl)
                ? plan.TargetServerRelativeUrl
                : plan.PreferredTargetServerRelativeUrl;
            var preferredUrl = string.IsNullOrWhiteSpace(plan.PreferredTargetWebUrl)
                ? AbsoluteUrl(plan.TargetSiteCollectionUrl, preferredPath)
                : plan.PreferredTargetWebUrl;

            var owned = inventory
                .Where(value => ExactShape(value, plan))
                .Where(value => string.Equals(value.OriginalIdentifier, plan.OriginalIdentifier, StringComparison.Ordinal))
                .Where(value => string.Equals(
                    value.MappingDigest,
                    ComputeDigestForTarget(plan, value.ServerRelativeUrl, value.Url),
                    StringComparison.OrdinalIgnoreCase))
                .OrderByDescending(value => PathEquals(value.ServerRelativeUrl, preferredPath))
                .ThenBy(value => value.ServerRelativeUrl, StringComparer.OrdinalIgnoreCase)
                .ThenBy(value => value.WebId)
                .FirstOrDefault();
            if (owned != null)
            {
                return Existing(plan, preferredPath, preferredUrl, owned, TopologyMaterializationDisposition.ReuseOwned,
                    "Reuse the deterministic migration-owned Web carrying the exact source identity and mapping digest.");
            }

            var exact = inventory.FirstOrDefault(value => PathEquals(value.ServerRelativeUrl, preferredPath));
            if (exact == null)
            {
                return New(preferredPath, preferredUrl, preferredPath, preferredUrl, false, null);
            }
            if (ExactShape(exact, plan)
                && string.IsNullOrWhiteSpace(exact.OriginalIdentifier)
                && string.IsNullOrWhiteSpace(exact.MappingDigest)
                && string.Equals(exact.Description, TopologyWebTargetInspector.InterruptedCreateDescription(plan), StringComparison.Ordinal))
            {
                return Existing(plan, preferredPath, preferredUrl, exact, TopologyMaterializationDisposition.RecoverInterruptedCreate,
                    "Recover the exact interrupted Web creation before applying ownership provenance.");
            }

            var targetPath = TopologyTargetPathAllocator.AllocateServerRelativePath(
                preferredPath,
                plan.OriginalIdentifier,
                inventory.Select(value => value.ServerRelativeUrl));
            var targetUrl = AbsoluteUrl(plan.TargetSiteCollectionUrl, targetPath);
            return New(
                preferredPath,
                preferredUrl,
                targetPath,
                targetUrl,
                true,
                "Allocated a stable suffix only at the Web node because the preferred target path is occupied by a foreign or incompatible Web.");
        }

        private static TopologyWebTargetPathResolution Existing(
            WebMappingPlan plan,
            string preferredPath,
            string preferredUrl,
            TopologyWebTargetInventoryItem existing,
            TopologyMaterializationDisposition disposition,
            string reason)
        {
            return new TopologyWebTargetPathResolution
            {
                PreferredTargetWebUrl = preferredUrl,
                PreferredTargetServerRelativeUrl = preferredPath,
                TargetWebUrl = existing.Url,
                TargetServerRelativeUrl = existing.ServerRelativeUrl,
                CollisionResolved = !PathEquals(existing.ServerRelativeUrl, preferredPath),
                Reason = reason,
                ExistingTarget = existing,
                ExistingDisposition = disposition
            };
        }

        private static TopologyWebTargetPathResolution New(
            string preferredPath,
            string preferredUrl,
            string targetPath,
            string targetUrl,
            bool collisionResolved,
            string reason)
        {
            return new TopologyWebTargetPathResolution
            {
                PreferredTargetWebUrl = preferredUrl,
                PreferredTargetServerRelativeUrl = preferredPath,
                TargetWebUrl = targetUrl,
                TargetServerRelativeUrl = targetPath,
                CollisionResolved = collisionResolved,
                Reason = reason,
                ExistingDisposition = TopologyMaterializationDisposition.CreateOwned
            };
        }

        private static bool ExactShape(TopologyWebTargetInventoryItem candidate, WebMappingPlan plan)
        {
            return string.Equals(candidate.Title, plan.TargetTitle, StringComparison.Ordinal)
                && TopologyWebTargetInspector.TemplateMatches(
                    candidate.Template,
                    candidate.Configuration,
                    plan.TargetTemplate,
                    plan.TargetConfiguration);
        }

        private static string ComputeDigestForTarget(WebMappingPlan plan, string path, string url)
        {
            var currentPath = plan.TargetServerRelativeUrl;
            var currentUrl = plan.TargetWebUrl;
            try
            {
                plan.TargetServerRelativeUrl = path;
                plan.TargetWebUrl = url;
                return TopologyPlanner.ComputeWebMappingDigest(plan);
            }
            finally
            {
                plan.TargetServerRelativeUrl = currentPath;
                plan.TargetWebUrl = currentUrl;
            }
        }

        private static string AbsoluteUrl(string siteUrl, string serverRelativePath)
        {
            return new Uri(new Uri(siteUrl).GetLeftPart(UriPartial.Authority) + serverRelativePath).AbsoluteUri.TrimEnd('/');
        }

        private static bool PathEquals(string left, string right)
        {
            return string.Equals(
                Uri.UnescapeDataString(left ?? string.Empty).TrimEnd('/'),
                Uri.UnescapeDataString(right ?? string.Empty).TrimEnd('/'),
                StringComparison.OrdinalIgnoreCase);
        }
    }
}
