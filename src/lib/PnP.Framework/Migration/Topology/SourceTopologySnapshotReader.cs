using Microsoft.SharePoint.Client;
using PnP.Framework.Migration.Evidence;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace PnP.Framework.Migration.Topology
{
    public static class SourceTopologySnapshotReader
    {
        public static SourceSiteCollectionSnapshot CaptureRequiredWebClosure(ClientContext context, IEnumerable<Guid> requiredWebIds)
        {
            if (context == null)
            {
                throw new ArgumentNullException(nameof(context));
            }
            if (requiredWebIds == null)
            {
                throw new ArgumentNullException(nameof(requiredWebIds));
            }

            var site = context.Site;
            var root = site.RootWeb;
            context.Load(site, value => value.Id, value => value.ServerRelativeUrl);
            LoadWeb(context, root);
            context.ExecuteQueryRetry();
            var requested = requiredWebIds.Where(value => value != Guid.Empty).Distinct().ToList();
            if (!requested.Contains(root.Id))
            {
                requested.Add(root.Id);
            }

            var captured = new Dictionary<Guid, SourceWebSnapshot>();
            var objects = new Dictionary<Guid, Web>();
            foreach (var webId in requested)
            {
                var web = webId == root.Id ? root : site.OpenWebById(webId);
                if (webId != root.Id)
                {
                    LoadWeb(context, web);
                }
                objects[webId] = web;
            }
            context.ExecuteQueryRetry();
            foreach (var pair in objects)
            {
                captured[pair.Key] = ToSnapshot(site.Id, root.Url, pair.Value, null);
            }

            while (captured.Values.Any(value => value.WebId != root.Id && !value.ParentWebId.HasValue))
            {
                var unresolved = captured.Values.Where(value => value.WebId != root.Id && !value.ParentWebId.HasValue).ToArray();
                var parents = new Dictionary<Guid, WebInformation>();
                foreach (var child in unresolved)
                {
                    var parent = objects[child.WebId].ParentWeb;
                    context.Load(parent, value => value.Id, value => value.ServerRelativeUrl, value => value.Title, value => value.WebTemplate, value => value.Configuration);
                    parents[child.WebId] = parent;
                }
                context.ExecuteQueryRetry();
                foreach (var child in unresolved)
                {
                    var parent = parents[child.WebId];
                    if (parent.Id == Guid.Empty)
                    {
                        throw new InvalidDataException("Source Web '" + child.WebUrl + "' did not resolve a direct parent Web.");
                    }
                    child.ParentWebId = parent.Id;
                    if (captured.ContainsKey(parent.Id))
                    {
                        continue;
                    }
                    var parentUrl = new Uri(new Uri(root.Url).GetLeftPart(UriPartial.Authority) + parent.ServerRelativeUrl).AbsoluteUri.TrimEnd('/');
                    captured[parent.Id] = new SourceWebSnapshot
                    {
                        SiteId = site.Id,
                        WebId = parent.Id,
                        SiteCollectionUrl = root.Url.TrimEnd('/'),
                        WebUrl = parentUrl,
                        ServerRelativeUrl = parent.ServerRelativeUrl,
                        Title = parent.Title,
                        WebTemplate = parent.WebTemplate,
                        Configuration = parent.Configuration
                    };
                    if (parent.Id != root.Id)
                    {
                        var parentWeb = site.OpenWebById(parent.Id);
                        objects[parent.Id] = parentWeb;
                    }
                }
            }

            return new SourceSiteCollectionSnapshot
            {
                SiteId = site.Id,
                SiteCollectionUrl = root.Url.TrimEnd('/'),
                ServerRelativeUrl = site.ServerRelativeUrl,
                RootWebId = root.Id,
                Webs = captured.Values.OrderBy(value => PathDepth(value.ServerRelativeUrl)).ThenBy(value => value.ServerRelativeUrl, StringComparer.OrdinalIgnoreCase).ToList(),
                Availability = EvidenceAvailability.Captured
            };
        }

        private static void LoadWeb(ClientContext context, Web web)
        {
            context.Load(web, value => value.Id, value => value.Url, value => value.ServerRelativeUrl, value => value.Title, value => value.WebTemplate, value => value.Configuration);
        }

        private static SourceWebSnapshot ToSnapshot(Guid siteId, string siteUrl, Web web, Guid? parentWebId)
        {
            return new SourceWebSnapshot
            {
                SiteId = siteId,
                WebId = web.Id,
                ParentWebId = parentWebId,
                SiteCollectionUrl = siteUrl.TrimEnd('/'),
                WebUrl = web.Url.TrimEnd('/'),
                ServerRelativeUrl = web.ServerRelativeUrl,
                Title = web.Title,
                WebTemplate = web.WebTemplate,
                Configuration = web.Configuration
            };
        }

        private static int PathDepth(string value)
        {
            return (value ?? string.Empty).Count(character => character == '/');
        }
    }
}
