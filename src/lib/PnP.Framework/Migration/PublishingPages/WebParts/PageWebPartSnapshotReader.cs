using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.WebParts;
using PnP.Framework.Migration.PublishingPages.Packaging;
using System;
using System.Collections.Generic;
using System.Linq;

namespace PnP.Framework.Migration.PublishingPages.WebParts
{
    internal static class PageWebPartSnapshotReader
    {
        public static List<PageWebPartSnapshot> Read(Web web, string pagePath, ICollection<string> blockers)
        {
            var result = new List<PageWebPartSnapshot>();
            var context = (ClientContext)web.Context;
            var file = web.GetFileByServerRelativePath(ResourcePath.FromDecodedUrl(pagePath));
            var manager = file.GetLimitedWebPartManager(PersonalizationScope.Shared);
            IEnumerable<WebPartDefinition> webParts;
            try
            {
                webParts = web.GetWebParts(pagePath).ToArray();
            }
            catch (ServerException exception)
            {
                blockers.Add($"Shared Web Part inventory could not be captured: {exception.Message}");
                return result;
            }

            foreach (var webPart in webParts)
            {
                string xml;
                try
                {
                    var export = manager.ExportWebPart(webPart.Id);
                    context.ExecuteQueryRetry();
                    xml = export.Value;
                }
                catch (ServerException exception)
                {
                    blockers.Add($"Web Part '{webPart.Id}' could not be exported: {exception.Message}");
                    continue;
                }

                if (string.IsNullOrWhiteSpace(xml))
                {
                    blockers.Add($"Web Part '{webPart.Id}' returned an empty export.");
                    continue;
                }

                var snapshot = new PageWebPartSnapshot
                {
                    Id = webPart.Id,
                    Title = webPart.WebPart.Title,
                    ZoneId = webPart.ZoneId,
                    ZoneIndex = webPart.WebPart.ZoneIndex,
                    Hidden = webPart.WebPart.Hidden,
                    ExportXml = xml,
                    ExportSha256 = PublishingPageDigest.ComputeSha256(xml)
                };
                result.Add(snapshot);

                var blocker = PageWebPartPortabilityPolicy.GetBlocker(xml);
                if (!string.IsNullOrWhiteSpace(blocker))
                {
                    var title = string.IsNullOrWhiteSpace(snapshot.Title) ? snapshot.Id.ToString() : snapshot.Title;
                    blockers.Add($"Web Part '{title}' ({snapshot.Id}) cannot be copied: {blocker}.");
                }
            }

            return result;
        }
    }
}
