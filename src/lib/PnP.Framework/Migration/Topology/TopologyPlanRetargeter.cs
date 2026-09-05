using System;
using System.Collections.Generic;
using System.Linq;

namespace PnP.Framework.Migration.Topology
{
    internal static class TopologyPlanRetargeter
    {
        public static void RetargetSiteCollection(
            SiteCollectionMappingPlan sitePlan,
            string targetSiteCollectionUrl,
            string reason)
        {
            if (sitePlan == null)
            {
                throw new ArgumentNullException(nameof(sitePlan));
            }
            if (!Uri.TryCreate(sitePlan.TargetSiteCollectionUrl, UriKind.Absolute, out var oldSiteUri)
                || !Uri.TryCreate(targetSiteCollectionUrl, UriKind.Absolute, out var newSiteUri)
                || oldSiteUri.Scheme != Uri.UriSchemeHttps
                || newSiteUri.Scheme != Uri.UriSchemeHttps
                || !string.Equals(oldSiteUri.Authority, newSiteUri.Authority, StringComparison.OrdinalIgnoreCase))
            {
                throw new ArgumentException("A target Site Collection URL on the same HTTPS tenant origin is required.", nameof(targetSiteCollectionUrl));
            }

            var oldSitePath = Normalize(oldSiteUri.AbsolutePath);
            var newSitePath = Normalize(newSiteUri.AbsolutePath);
            var preferredUrl = sitePlan.PreferredTargetSiteCollectionUrl ?? sitePlan.TargetSiteCollectionUrl;
            sitePlan.PreferredTargetSiteCollectionUrl = preferredUrl;
            sitePlan.TargetSiteCollectionUrl = newSiteUri.AbsoluteUri.TrimEnd('/');
            sitePlan.TargetSiteCollisionResolved = !string.Equals(
                new Uri(preferredUrl).AbsoluteUri.TrimEnd('/'),
                sitePlan.TargetSiteCollectionUrl,
                StringComparison.OrdinalIgnoreCase);
            sitePlan.TargetSiteResolutionReason = sitePlan.TargetSiteCollisionResolved ? reason : null;
            sitePlan.ExpectedTargetSiteId = null;

            var byId = sitePlan.Webs.ToDictionary(value => value.SourceWebId);
            foreach (var web in sitePlan.Webs.OrderBy(value => Depth(value.TargetServerRelativeUrl)))
            {
                var oldPath = Normalize(web.TargetServerRelativeUrl);
                if (!string.Equals(oldPath, oldSitePath, StringComparison.OrdinalIgnoreCase)
                    && !oldPath.StartsWith(oldSitePath + "/", StringComparison.OrdinalIgnoreCase))
                {
                    throw new InvalidOperationException("A mapped Web is outside its target Site Collection.");
                }

                web.TargetSiteCollectionUrl = sitePlan.TargetSiteCollectionUrl;
                web.TargetServerRelativeUrl = newSitePath + oldPath.Substring(oldSitePath.Length);
                web.TargetWebUrl = AbsoluteUrl(sitePlan.TargetSiteCollectionUrl, web.TargetServerRelativeUrl);
                web.TargetParentWebUrl = web.SourceParentWebId.HasValue
                    ? byId[web.SourceParentWebId.Value].TargetWebUrl
                    : null;
            }
        }

        public static void RetargetWeb(
            SiteCollectionMappingPlan sitePlan,
            Guid sourceWebId,
            string targetServerRelativeUrl)
        {
            if (sitePlan == null)
            {
                throw new ArgumentNullException(nameof(sitePlan));
            }

            var byId = sitePlan.Webs.ToDictionary(value => value.SourceWebId);
            WebMappingPlan selected;
            if (!byId.TryGetValue(sourceWebId, out selected) || selected.Kind != TopologyNodeKind.ChildWeb)
            {
                throw new ArgumentException("A mapped child Web is required.", nameof(sourceWebId));
            }

            var newPath = Normalize(targetServerRelativeUrl);
            WebMappingPlan parent;
            if (!selected.SourceParentWebId.HasValue || !byId.TryGetValue(selected.SourceParentWebId.Value, out parent)
                || (!string.Equals(newPath, parent.TargetServerRelativeUrl, StringComparison.OrdinalIgnoreCase)
                    && !newPath.StartsWith(parent.TargetServerRelativeUrl.TrimEnd('/') + "/", StringComparison.OrdinalIgnoreCase)))
            {
                throw new ArgumentException("The retargeted child Web path must remain below its mapped source parent.", nameof(targetServerRelativeUrl));
            }

            var oldPaths = sitePlan.Webs.ToDictionary(value => value.SourceWebId, value => value.TargetServerRelativeUrl);
            var affected = DescendantsInclusive(sourceWebId, sitePlan.Webs)
                .OrderBy(value => Depth(oldPaths[value.SourceWebId]))
                .ToArray();
            foreach (var web in affected)
            {
                if (web.SourceWebId == sourceWebId)
                {
                    web.TargetServerRelativeUrl = newPath;
                }
                else
                {
                    var oldParentPath = oldPaths[web.SourceParentWebId.Value].TrimEnd('/');
                    var oldPath = oldPaths[web.SourceWebId];
                    if (!oldPath.StartsWith(oldParentPath + "/", StringComparison.OrdinalIgnoreCase))
                    {
                        throw new InvalidOperationException("A descendant target path is outside its mapped source parent.");
                    }
                    web.TargetServerRelativeUrl = byId[web.SourceParentWebId.Value].TargetServerRelativeUrl.TrimEnd('/')
                        + oldPath.Substring(oldParentPath.Length);
                }
                web.TargetWebUrl = AbsoluteUrl(sitePlan.TargetSiteCollectionUrl, web.TargetServerRelativeUrl);
                web.TargetParentWebUrl = web.SourceParentWebId.HasValue
                    ? byId[web.SourceParentWebId.Value].TargetWebUrl
                    : null;
            }
        }

        private static IEnumerable<WebMappingPlan> DescendantsInclusive(Guid sourceWebId, IEnumerable<WebMappingPlan> webs)
        {
            var all = webs.ToArray();
            var selected = all.Single(value => value.SourceWebId == sourceWebId);
            yield return selected;
            var queue = new Queue<Guid>();
            queue.Enqueue(sourceWebId);
            while (queue.Count > 0)
            {
                var parent = queue.Dequeue();
                foreach (var child in all.Where(value => value.SourceParentWebId == parent).OrderBy(value => value.SourceWebId))
                {
                    yield return child;
                    queue.Enqueue(child.SourceWebId);
                }
            }
        }

        private static string AbsoluteUrl(string siteUrl, string serverRelativePath)
        {
            return new Uri(new Uri(siteUrl).GetLeftPart(UriPartial.Authority) + serverRelativePath).AbsoluteUri.TrimEnd('/');
        }

        private static string Normalize(string value)
        {
            if (string.IsNullOrWhiteSpace(value) || !value.StartsWith("/", StringComparison.Ordinal))
            {
                throw new ArgumentException("A server-relative target Web path is required.", nameof(value));
            }
            return Uri.UnescapeDataString(value).TrimEnd('/');
        }

        private static int Depth(string value)
        {
            return (value ?? string.Empty).Count(character => character == '/');
        }
    }
}
