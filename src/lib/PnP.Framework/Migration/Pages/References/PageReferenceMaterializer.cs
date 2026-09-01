using Microsoft.SharePoint.Client;
using PnP.Framework.Migration.Pages.Packaging;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace PnP.Framework.Migration.Pages.References
{
    internal static class PageReferenceMaterializer
    {
        public static int Materialize(
            ClientContext context,
            IEnumerable<PageReferenceSnapshot> snapshots,
            IEnumerable<PageReferenceAction> actions)
        {
            var web = context.Web;
            web.EnsureProperty(value => value.ServerRelativeUrl);
            var snapshotById = snapshots.ToDictionary(item => item.Id, StringComparer.Ordinal);
            var count = 0;
            foreach (var action in actions
                         .Where(item => item.Disposition == PageReferenceDisposition.MaterializeAtTarget)
                         .GroupBy(item => item.TargetServerRelativeUrl, StringComparer.OrdinalIgnoreCase)
                         .Select(group => group.First()))
            {
                if (!snapshotById.TryGetValue(action.SnapshotDependencyId, out var snapshot))
                {
                    throw new InvalidDataException($"Dependency action references missing snapshot '{action.SnapshotDependencyId}'.");
                }

                var bytes = Convert.FromBase64String(snapshot.ContentBase64 ?? string.Empty);
                if (!string.Equals(
                        PageDigest.ComputeSha256(bytes),
                        snapshot.ContentSha256,
                        StringComparison.OrdinalIgnoreCase))
                {
                    throw new InvalidDataException($"Dependency payload digest mismatch: {action.TargetServerRelativeUrl}");
                }

                if (!PagePath.IsWithin(action.TargetServerRelativeUrl, web.ServerRelativeUrl))
                {
                    throw new InvalidOperationException($"Dependency target escapes the target web boundary: {action.TargetServerRelativeUrl}");
                }

                var webPath = web.ServerRelativeUrl.TrimEnd('/');
                var relativePath = action.TargetServerRelativeUrl.Substring(webPath.Length).TrimStart('/');
                var separator = relativePath.LastIndexOf('/');
                var folderPath = separator < 0 ? string.Empty : relativePath.Substring(0, separator);
                var fileName = separator < 0 ? relativePath : relativePath.Substring(separator + 1);
                var folder = string.IsNullOrEmpty(folderPath)
                    ? web.RootFolder
                    : web.EnsureFolderPath(folderPath);
                using (var stream = new MemoryStream(bytes, writable: false))
                {
                    var created = folder.UploadFile(fileName, stream, false);
                    created.EnsureProperties(file => file.Exists, file => file.ServerRelativeUrl);
                    if (!created.Exists)
                    {
                        throw new InvalidOperationException($"SharePoint did not persist dependency '{action.TargetServerRelativeUrl}'.");
                    }
                }

                count++;
            }

            return count;
        }
    }
}
