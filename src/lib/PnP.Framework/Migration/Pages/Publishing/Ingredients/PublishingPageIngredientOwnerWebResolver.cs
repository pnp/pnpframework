using PnP.Framework.Migration.Pages.Publishing.Capture;
using PnP.Framework.Migration.Topology;
using System;
using System.Linq;

namespace PnP.Framework.Migration.Pages.Publishing.Ingredients
{
    /// <summary>
    /// Resolves the captured Web transaction that owns a schema or asset. The
    /// graph stores source identities; target retargeting remains sealed in the
    /// topology plan.
    /// </summary>
    internal static class PublishingPageIngredientOwnerWebResolver
    {
        public static string Root(PublishingPageCaptureBundle snapshot)
        {
            var topology = snapshot?.SourceTopology;
            if (topology == null || topology.RootWebId == Guid.Empty)
            {
                return null;
            }

            var root = topology.Webs?.FirstOrDefault(value => value != null
                && value.WebId == topology.RootWebId);
            return root == null
                ? null
                : PublishingPageIngredientIds.Web(root.SiteId, root.WebId);
        }

        public static string ExactOrContaining(PublishingPageCaptureBundle snapshot, string sourceUrlOrPath)
        {
            var topology = snapshot?.SourceTopology;
            var path = Path(sourceUrlOrPath);
            if (topology?.Webs == null || string.IsNullOrWhiteSpace(path))
            {
                return null;
            }

            var owner = topology.Webs
                .Where(value => value != null)
                .Select(value => new { Web = value, Path = Path(value.ServerRelativeUrl ?? value.WebUrl) })
                .Where(value => Contains(path, value.Path))
                .OrderByDescending(value => value.Path.Length)
                .ThenBy(value => value.Web.WebId)
                .Select(value => value.Web)
                .FirstOrDefault();
            return owner == null
                ? null
                : PublishingPageIngredientIds.Web(owner.SiteId, owner.WebId);
        }

        private static bool Contains(string candidatePath, string webPath)
        {
            if (string.IsNullOrWhiteSpace(candidatePath) || string.IsNullOrWhiteSpace(webPath))
            {
                return false;
            }
            if (string.Equals(webPath, "/", StringComparison.Ordinal))
            {
                return candidatePath.StartsWith("/", StringComparison.Ordinal);
            }
            return string.Equals(candidatePath, webPath, StringComparison.OrdinalIgnoreCase)
                || candidatePath.StartsWith(webPath + "/", StringComparison.OrdinalIgnoreCase);
        }

        private static string Path(string value)
        {
            if (string.IsNullOrWhiteSpace(value))
            {
                return null;
            }
            var path = Uri.TryCreate(value, UriKind.Absolute, out var absolute)
                ? absolute.AbsolutePath
                : value;
            path = Uri.UnescapeDataString(path).Replace('\\', '/').Trim();
            if (!path.StartsWith("/", StringComparison.Ordinal))
            {
                path = "/" + path;
            }
            path = path.TrimEnd('/');
            return path.Length == 0 ? "/" : path;
        }
    }
}
