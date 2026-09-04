using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace PnP.Framework.Migration.Topology
{
    /// <summary>
    /// Validates the sealed Site/Web transaction graph before any target inspection or mutation.
    /// A child Web must name one captured direct parent and one URL segment below that parent.
    /// </summary>
    public static class TopologyPlanValidator
    {
        public static void Validate(TopologyPlan plan, bool requireDigest = true)
        {
            if (plan == null)
            {
                throw new ArgumentNullException(nameof(plan));
            }

            var errors = new List<string>();
            if (!string.Equals(plan.SchemaVersion, "pnp-topology-plan/v1", StringComparison.Ordinal))
            {
                errors.Add("Unsupported topology plan schema.");
            }
            if (requireDigest && (string.IsNullOrWhiteSpace(plan.PlanDigest)
                || !string.Equals(plan.PlanDigest, TopologyPlanner.ComputeDigest(plan), StringComparison.OrdinalIgnoreCase)))
            {
                errors.Add("The topology plan digest is absent or invalid.");
            }

            var sites = plan.SiteCollections ?? new List<SiteCollectionMappingPlan>();
            AddDuplicates(errors, sites.Select(value => value.SourceSiteId), "source Site ID");
            AddDuplicates(errors, sites.Select(value => NormalizeUrl(value.TargetSiteCollectionUrl)), "target Site URL", StringComparer.OrdinalIgnoreCase);
            foreach (var site in sites)
            {
                ValidateSite(site, errors);
            }

            if (errors.Count > 0)
            {
                throw new InvalidDataException("Invalid topology plan: " + string.Join(" ", errors));
            }
        }

        private static void ValidateSite(SiteCollectionMappingPlan site, ICollection<string> errors)
        {
            if (site == null || site.SourceSiteId == Guid.Empty
                || !IsHttps(site.SourceSiteCollectionUrl)
                || !IsHttps(site.TargetSiteCollectionUrl)
                || !string.Equals(site.OriginalIdentifier, TopologyPlanner.SiteOriginalIdentifier(site.SourceSiteId), StringComparison.Ordinal))
            {
                errors.Add("A Site mapping has invalid identity, URL, or provenance.");
                return;
            }

            var webs = site.Webs ?? new List<WebMappingPlan>();
            AddDuplicates(errors, webs.Select(value => value.SourceWebId), "source Web ID");
            AddDuplicates(errors, webs.Select(value => NormalizeUrl(value.TargetWebUrl)), "target Web URL", StringComparer.OrdinalIgnoreCase);
            AddDuplicates(errors, webs.Select(value => NormalizePath(value.TargetServerRelativeUrl)), "target Web path", StringComparer.OrdinalIgnoreCase);
            var roots = webs.Where(value => value.Kind == TopologyNodeKind.SiteCollectionRoot).ToArray();
            if (roots.Length != 1)
            {
                errors.Add("Source Site '" + site.SourceSiteId.ToString("D") + "' must contain exactly one Site Collection root Web.");
                return;
            }

            var byId = webs.ToDictionary(value => value.SourceWebId);
            var root = roots[0];
            if (root.SourceParentWebId.HasValue
                || !UrlEquals(root.TargetWebUrl, site.TargetSiteCollectionUrl)
                || !string.IsNullOrWhiteSpace(root.TargetParentWebUrl))
            {
                errors.Add("The root Web mapping for source Site '" + site.SourceSiteId.ToString("D") + "' is inconsistent with its target Site URL.");
            }

            foreach (var web in webs)
            {
                if (web.SourceSiteId != site.SourceSiteId
                    || web.SourceWebId == Guid.Empty
                    || !string.Equals(web.OriginalIdentifier, TopologyPlanner.WebOriginalIdentifier(site.SourceSiteId, web.SourceWebId), StringComparison.Ordinal)
                    || !IsHttps(web.SourceWebUrl)
                    || !IsHttps(web.TargetWebUrl)
                    || !UrlEquals(web.TargetSiteCollectionUrl, site.TargetSiteCollectionUrl)
                    || !PathEquals(new Uri(web.SourceWebUrl).AbsolutePath, web.SourceServerRelativeUrl)
                    || !PathEquals(new Uri(web.TargetWebUrl).AbsolutePath, web.TargetServerRelativeUrl))
                {
                    errors.Add("Web '" + web.SourceWebId.ToString("D") + "' has invalid identity, URL, or provenance.");
                    continue;
                }

                if (web.Kind == TopologyNodeKind.SiteCollectionRoot)
                {
                    continue;
                }
                if (web.Kind != TopologyNodeKind.ChildWeb
                    || !web.SourceParentWebId.HasValue
                    || !byId.TryGetValue(web.SourceParentWebId.Value, out var parent))
                {
                    errors.Add("Child Web '" + web.SourceWebId.ToString("D") + "' has no captured direct parent mapping.");
                    continue;
                }
                if (!UrlEquals(web.TargetParentWebUrl, parent.TargetWebUrl)
                    || !IsOneDirectSegmentBelow(web.TargetServerRelativeUrl, parent.TargetServerRelativeUrl)
                    || !IsOneDirectSegmentBelow(web.SourceServerRelativeUrl, parent.SourceServerRelativeUrl))
                {
                    errors.Add("Child Web '" + web.SourceWebId.ToString("D") + "' is not one direct URL segment below its captured parent.");
                }

                var visited = new HashSet<Guid> { web.SourceWebId };
                var cursor = web;
                while (cursor.SourceParentWebId.HasValue)
                {
                    if (!visited.Add(cursor.SourceParentWebId.Value)
                        || !byId.TryGetValue(cursor.SourceParentWebId.Value, out cursor))
                    {
                        errors.Add("Child Web '" + web.SourceWebId.ToString("D") + "' has a cyclic or incomplete parent chain.");
                        break;
                    }
                }
            }
        }

        private static bool IsOneDirectSegmentBelow(string childValue, string parentValue)
        {
            var child = NormalizePath(childValue);
            var parent = NormalizePath(parentValue).TrimEnd('/');
            if (!child.StartsWith(parent + "/", StringComparison.OrdinalIgnoreCase))
            {
                return false;
            }
            var relative = Uri.UnescapeDataString(child.Substring(parent.Length + 1));
            return !string.IsNullOrWhiteSpace(relative)
                && relative.IndexOf('/') < 0
                && relative.IndexOf('\\') < 0
                && relative != "."
                && relative != "..";
        }

        private static bool IsHttps(string value)
        {
            return Uri.TryCreate(value, UriKind.Absolute, out var uri) && uri.Scheme == Uri.UriSchemeHttps;
        }

        private static string NormalizeUrl(string value)
        {
            return Uri.TryCreate(value, UriKind.Absolute, out var uri)
                ? uri.AbsoluteUri.TrimEnd('/')
                : "invalid:" + (value ?? string.Empty);
        }

        private static string NormalizePath(string value)
        {
            if (string.IsNullOrWhiteSpace(value))
            {
                return string.Empty;
            }
            return Uri.UnescapeDataString(value).Replace('\\', '/').TrimEnd('/');
        }

        private static bool UrlEquals(string left, string right)
        {
            return string.Equals(NormalizeUrl(left), NormalizeUrl(right), StringComparison.OrdinalIgnoreCase);
        }

        private static bool PathEquals(string left, string right)
        {
            return string.Equals(NormalizePath(left), NormalizePath(right), StringComparison.OrdinalIgnoreCase);
        }

        private static void AddDuplicates<T>(
            ICollection<string> errors,
            IEnumerable<T> values,
            string description,
            IEqualityComparer<T> comparer = null)
        {
            var duplicate = values.GroupBy(value => value, comparer ?? EqualityComparer<T>.Default)
                .FirstOrDefault(value => value.Count() > 1);
            if (duplicate != null)
            {
                errors.Add("Duplicate " + description + " '" + duplicate.Key + "'.");
            }
        }
    }
}
