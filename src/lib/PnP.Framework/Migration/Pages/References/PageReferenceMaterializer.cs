using Microsoft.SharePoint.Client;
using PnP.Framework.Migration.Pages.Packaging;
using PnP.Framework.Migration.Packaging;
using PnP.Framework.Migration.Topology;
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
            IEnumerable<PageReferenceAction> actions,
            TopologyMaterializationReceipt topology)
        {
            var pageWeb = context.Web;
            var rootWeb = context.Site.RootWeb;
            context.Load(pageWeb, value => value.Id, value => value.ServerRelativeUrl);
            context.Load(rootWeb, value => value.Id, value => value.ServerRelativeUrl);
            context.ExecuteQueryRetry();
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

                var owner = ResolveOwner(
                    context,
                    pageWeb,
                    rootWeb,
                    action.TargetServerRelativeUrl,
                    topology);
                var webPath = owner.ServerRelativeUrl == "/"
                    ? string.Empty
                    : owner.ServerRelativeUrl.TrimEnd('/');
                var relativePath = action.TargetServerRelativeUrl.Substring(webPath.Length).TrimStart('/');
                var separator = relativePath.LastIndexOf('/');
                var folderPath = separator < 0 ? string.Empty : relativePath.Substring(0, separator);
                var fileName = separator < 0 ? relativePath : relativePath.Substring(separator + 1);
                var folder = string.IsNullOrEmpty(folderPath)
                    ? owner.Web.RootFolder
                    : owner.Web.EnsureFolderPath(folderPath);
                var targetFile = owner.Web.GetFileByServerRelativePath(
                    ResourcePath.FromDecodedUrl(action.TargetServerRelativeUrl));
                var existingStream = targetFile.OpenBinaryStream();
                context.Load(targetFile, value => value.Exists);
                var exists = false;
                try
                {
                    context.ExecuteQueryRetry();
                    exists = targetFile.Exists && existingStream.Value != null;
                }
                catch (ServerException exception) when (IsFileNotFound(exception))
                {
                }
                if (exists)
                {
                    using (existingStream.Value)
                    {
                        VerifyDigest(existingStream.Value, snapshot, action.TargetServerRelativeUrl);
                    }
                }
                else
                {
                    using (var stream = new MemoryStream(bytes, writable: false))
                    {
                        var created = folder.UploadFile(fileName, stream, false);
                        created.EnsureProperties(file => file.Exists, file => file.ServerRelativeUrl);
                        if (!created.Exists)
                        {
                            throw new InvalidOperationException($"SharePoint did not persist dependency '{action.TargetServerRelativeUrl}'.");
                        }
                    }
                }

                var readback = owner.Web.GetFileByServerRelativePath(
                    ResourcePath.FromDecodedUrl(action.TargetServerRelativeUrl));
                var readbackStream = readback.OpenBinaryStream();
                context.Load(readback, value => value.Exists);
                context.ExecuteQueryRetry();
                if (!readback.Exists || readbackStream.Value == null)
                {
                    throw new InvalidOperationException($"Fresh dependency readback failed: {action.TargetServerRelativeUrl}");
                }
                using (readbackStream.Value)
                {
                    VerifyDigest(readbackStream.Value, snapshot, action.TargetServerRelativeUrl);
                }

                count++;
            }

            return count;
        }

        private static DependencyOwner ResolveOwner(
            ClientContext context,
            Web pageWeb,
            Web rootWeb,
            string targetPath,
            TopologyMaterializationReceipt topology)
        {
            var candidate = topology?.Webs
                .Where(value => value != null
                    && value.TargetWebId != Guid.Empty
                    && !string.IsNullOrWhiteSpace(value.TargetWebUrl)
                    && PagePath.IsWithin(
                        targetPath,
                        Uri.UnescapeDataString(new Uri(value.TargetWebUrl).AbsolutePath)))
                .OrderByDescending(value => new Uri(value.TargetWebUrl).AbsolutePath.Length)
                .FirstOrDefault();
            if (candidate == null)
            {
                throw new InvalidOperationException(
                    $"Dependency target is outside the reviewed target topology: {targetPath}");
            }

            var ownerWeb = candidate.TargetWebId == pageWeb.Id
                ? pageWeb
                : candidate.TargetWebId == rootWeb.Id
                    ? rootWeb
                    : context.Site.OpenWebById(candidate.TargetWebId);
            return new DependencyOwner
            {
                Web = ownerWeb,
                ServerRelativeUrl = Uri.UnescapeDataString(new Uri(candidate.TargetWebUrl).AbsolutePath)
            };
        }

        private static void VerifyDigest(
            Stream stream,
            PageReferenceSnapshot snapshot,
            string targetPath)
        {
            if (stream == null
                || !string.Equals(MigrationDigest.ComputeSha256(stream), snapshot.ContentSha256, StringComparison.OrdinalIgnoreCase))
            {
                throw new InvalidOperationException($"Target dependency digest differs: {targetPath}");
            }
        }

        private static bool IsFileNotFound(ServerException exception)
        {
            return string.Equals(exception.ServerErrorTypeName, "System.IO.FileNotFoundException", StringComparison.Ordinal)
                || exception.ServerErrorCode == -2147024894;
        }

        private sealed class DependencyOwner
        {
            public Web Web { get; set; }

            public string ServerRelativeUrl { get; set; }
        }
    }
}
