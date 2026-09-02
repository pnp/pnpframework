using Microsoft.SharePoint.Client;
using System;

namespace PnP.Framework.Migration.Topology
{
    internal sealed class TopologyTargetInspectionScope
    {
        private TopologyTargetInspectionScope(
            ClientContext anchorContext,
            Guid approvedSiteId,
            Guid approvedHostId,
            string approvedHostUrl,
            LoadedRootTarget loadedRoot)
        {
            AnchorContext = anchorContext;
            ApprovedSiteId = approvedSiteId;
            ApprovedHostId = approvedHostId;
            ApprovedHostUrl = approvedHostUrl;
            LoadedRoot = loadedRoot;
        }

        public ClientContext AnchorContext { get; }
        public Guid ApprovedSiteId { get; }
        public Guid ApprovedHostId { get; }
        public string ApprovedHostUrl { get; }
        public LoadedRootTarget LoadedRoot { get; }

        public static TopologyTargetInspectionScope Create(ClientContext anchorContext, string approvedHostWebUrl)
        {
            if (!anchorContext.Site.IsPropertyAvailable("Id")
                || !anchorContext.Web.IsPropertyAvailable("Id")
                || !anchorContext.Web.IsPropertyAvailable("Url"))
            {
                anchorContext.Load(anchorContext.Site, value => value.Id);
                anchorContext.Load(anchorContext.Web, value => value.Id, value => value.Url);
                anchorContext.ExecuteQueryRetry();
            }

            var approvedSiteId = anchorContext.Site.Id;
            return new TopologyTargetInspectionScope(
                anchorContext,
                approvedSiteId,
                anchorContext.Web.Id,
                TargetUrl.NormalizeAbsolute(approvedHostWebUrl ?? anchorContext.Web.Url),
                TryGetLoadedRoot(anchorContext, approvedSiteId));
        }

        private static LoadedRootTarget TryGetLoadedRoot(ClientContext context, Guid siteId)
        {
            var root = context.Site.RootWeb;
            if (!root.IsPropertyAvailable("Id")
                || !root.IsPropertyAvailable("Url")
                || !root.IsPropertyAvailable("Title")
                || !root.IsPropertyAvailable("WebTemplate")
                || !root.IsPropertyAvailable("Configuration"))
            {
                return null;
            }

            return new LoadedRootTarget(
                siteId,
                root.Id,
                root.Url,
                root.Title,
                root.WebTemplate,
                root.Configuration);
        }

        internal sealed class LoadedRootTarget
        {
            public LoadedRootTarget(Guid siteId, Guid webId, string url, string title, string template, int configuration)
            {
                SiteId = siteId;
                WebId = webId;
                Url = url;
                Title = title;
                Template = template;
                Configuration = configuration;
            }

            public Guid SiteId { get; }
            public Guid WebId { get; }
            public string Url { get; }
            public string Title { get; }
            public string Template { get; }
            public int Configuration { get; }
        }

        internal static class TargetUrl
        {
            public static bool Equals(string left, string right)
            {
                return string.Equals(NormalizeAbsolute(left), NormalizeAbsolute(right), StringComparison.OrdinalIgnoreCase);
            }

            public static string NormalizeAbsolute(string value)
            {
                return new Uri(value).AbsoluteUri.TrimEnd('/');
            }

            public static string NormalizePath(string value)
            {
                return Uri.UnescapeDataString(value ?? string.Empty).TrimEnd('/');
            }

            public static int Depth(string path)
            {
                return NormalizePath(path).Split(new[] { '/' }, StringSplitOptions.RemoveEmptyEntries).Length;
            }
        }
    }
}
