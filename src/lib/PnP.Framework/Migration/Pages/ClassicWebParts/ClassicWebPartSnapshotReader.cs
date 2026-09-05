using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.WebParts;
using PnP.Framework.Migration.Pages.Packaging;
using System;
using System.Collections.Generic;
using System.Linq;

namespace PnP.Framework.Migration.Pages.ClassicWebParts
{
    internal static class ClassicWebPartSnapshotReader
    {
        public static List<ClassicWebPartSnapshot> Read(Web web, string pagePath, ICollection<string> blockers)
        {
            var result = new List<ClassicWebPartSnapshot>();
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

                var snapshot = new ClassicWebPartSnapshot
                {
                    Id = webPart.Id,
                    Title = webPart.WebPart.Title,
                    TypeName = ClassicWebPartMetadataParser.ReadTypeName(xml),
                    ZoneId = webPart.ZoneId,
                    ZoneIndex = webPart.WebPart.ZoneIndex,
                    Hidden = webPart.WebPart.Hidden,
                    ExportXml = xml,
                    ExportSha256 = PageDigest.ComputeSha256(xml)
                };
                result.Add(snapshot);
            }

            return result;
        }
    }
}
