using Microsoft.SharePoint.Client;
using PnP.Framework.Migration.Lists.Planning;
using PnP.Framework.Migration.Lists.Views;
using PnP.Framework.Migration.Packaging;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace PnP.Framework.Migration.Lists.Execution
{
    internal static class ListViewRenderingResourceMaterializer
    {
        public static int Ensure(
            ClientContext context,
            ListMaterializationPlan plan,
            IMigrationArtifactStore artifactStore)
        {
            var created = 0;
            foreach (var resource in plan.ViewRenderingResources
                         .Where(value => value.Disposition == ListViewRenderingResourceMaterializationDisposition.CreateOrReuseExact)
                         .GroupBy(value => value.TargetServerRelativeUrl, StringComparer.OrdinalIgnoreCase)
                         .Select(group => group.First()))
            {
                if (resource.SourceArtifact == null || string.IsNullOrWhiteSpace(resource.TargetServerRelativeUrl))
                {
                    throw new InvalidOperationException("View rendering-resource plan is incomplete: " + resource.SourceResourceId);
                }
                var bytes = MigrationArtifact.ReadAllBytes(
                    resource.SourceArtifact,
                    resource.SourceContentBase64,
                    artifactStore);
                if (EnsureResource(context, resource, bytes))
                {
                    created++;
                }
            }
            return created;
        }

        public static string RewriteJsLink(
            string sourceValue,
            ListViewSnapshot sourceView,
            IEnumerable<ListViewRenderingResourceMaterializationPlan> resources)
        {
            if (string.IsNullOrWhiteSpace(sourceValue) || sourceView == null)
            {
                return sourceValue;
            }
            var plans = (resources ?? Array.Empty<ListViewRenderingResourceMaterializationPlan>())
                .Where(value => value != null && !string.IsNullOrWhiteSpace(value.SourceResourceId))
                .GroupBy(value => value.SourceResourceId, StringComparer.Ordinal)
                .ToDictionary(group => group.Key, group => group.First(), StringComparer.Ordinal);
            var rewrites = (sourceView.RenderingResourceBindings ?? Array.Empty<ListViewRenderingResourceBindingSnapshot>())
                .Where(value => value != null
                    && string.Equals(value.SourceProperty, "JSLink", StringComparison.OrdinalIgnoreCase)
                    && plans.ContainsKey(value.ResourceId ?? string.Empty))
                .ToDictionary(
                    value => value.OriginalReference,
                    value => TargetAuthoredReference(value.OriginalReference, plans[value.ResourceId]),
                    StringComparer.OrdinalIgnoreCase);
            return string.Join("|", sourceValue.Split('|').Select(value =>
            {
                var trimmed = value.Trim();
                return rewrites.TryGetValue(trimmed, out var target) ? target : trimmed;
            }));
        }

        public static int Verify(
            ClientContext context,
            ListMaterializationPlan plan,
            ICollection<string> diagnostics)
        {
            var pageWeb = context.Web;
            var rootWeb = context.Site.RootWeb;
            context.Load(pageWeb, value => value.Id, value => value.ServerRelativeUrl);
            context.Load(rootWeb, value => value.Id, value => value.ServerRelativeUrl);
            context.ExecuteQueryRetry();
            var verified = 0;
            foreach (var resource in plan.ViewRenderingResources
                         .Where(value => value.Disposition == ListViewRenderingResourceMaterializationDisposition.CreateOrReuseExact)
                         .GroupBy(value => value.TargetServerRelativeUrl, StringComparer.OrdinalIgnoreCase)
                         .Select(group => group.First()))
            {
                try
                {
                    var owner = ResolveOwner(pageWeb, rootWeb, resource.TargetServerRelativeUrl);
                    var file = owner.GetFileByServerRelativePath(ResourcePath.FromDecodedUrl(resource.TargetServerRelativeUrl));
                    var stream = file.OpenBinaryStream();
                    context.Load(file, value => value.Exists);
                    context.ExecuteQueryRetry();
                    if (!file.Exists || stream.Value == null)
                    {
                        diagnostics.Add("Target View rendering resource is missing: " + resource.TargetServerRelativeUrl + ".");
                        continue;
                    }
                    using (stream.Value)
                    {
                        var digest = MigrationDigest.ComputeSha256(stream.Value);
                        if (!string.Equals(digest, resource.SourceArtifact?.Sha256, StringComparison.OrdinalIgnoreCase))
                        {
                            diagnostics.Add("Target View rendering-resource digest differs: " + resource.TargetServerRelativeUrl + ".");
                            continue;
                        }
                    }
                    verified++;
                }
                catch (ServerException exception) when (IsFileNotFound(exception))
                {
                    diagnostics.Add("Target View rendering resource is missing: " + resource.TargetServerRelativeUrl + ".");
                }
            }
            return verified;
        }

        private static bool EnsureResource(
            ClientContext context,
            ListViewRenderingResourceMaterializationPlan plan,
            byte[] bytes)
        {
            var pageWeb = context.Web;
            var rootWeb = context.Site.RootWeb;
            context.Load(pageWeb, value => value.Id, value => value.ServerRelativeUrl);
            context.Load(rootWeb, value => value.Id, value => value.ServerRelativeUrl);
            context.ExecuteQueryRetry();
            var owner = ResolveOwner(pageWeb, rootWeb, plan.TargetServerRelativeUrl);
            var file = owner.GetFileByServerRelativePath(ResourcePath.FromDecodedUrl(plan.TargetServerRelativeUrl));
            var stream = file.OpenBinaryStream();
            context.Load(file, value => value.Exists);
            try
            {
                context.ExecuteQueryRetry();
                if (file.Exists && stream.Value != null)
                {
                    using (stream.Value)
                    {
                        var existingDigest = MigrationDigest.ComputeSha256(stream.Value);
                        if (!string.Equals(existingDigest, plan.SourceArtifact.Sha256, StringComparison.OrdinalIgnoreCase))
                        {
                            throw new InvalidOperationException("Target View rendering-resource collision: " + plan.TargetServerRelativeUrl);
                        }
                    }
                    return false;
                }
            }
            catch (ServerException exception) when (IsFileNotFound(exception))
            {
            }

            var relative = RelativeToWeb(owner.ServerRelativeUrl, plan.TargetServerRelativeUrl);
            var separator = relative.LastIndexOf('/');
            var folderPath = separator < 0 ? string.Empty : relative.Substring(0, separator);
            var fileName = separator < 0 ? relative : relative.Substring(separator + 1);
            ValidateAssetFolder(folderPath);
            var folder = EnsureAssetFolder(context, owner, folderPath);
            using (var content = new MemoryStream(bytes, false))
            {
                var created = folder.UploadFile(fileName, content, false);
                created.EnsureProperties(value => value.Exists, value => value.ServerRelativeUrl);
                if (!created.Exists)
                {
                    throw new InvalidOperationException("SharePoint did not persist View rendering resource '" + plan.TargetServerRelativeUrl + "'.");
                }
            }

            var readback = owner.GetFileByServerRelativePath(ResourcePath.FromDecodedUrl(plan.TargetServerRelativeUrl));
            var readbackStream = readback.OpenBinaryStream();
            context.Load(readback, value => value.Exists);
            context.ExecuteQueryRetry();
            if (!readback.Exists || readbackStream.Value == null)
            {
                throw new InvalidOperationException("Fresh View rendering-resource readback failed: " + plan.TargetServerRelativeUrl);
            }
            using (readbackStream.Value)
            {
                var digest = MigrationDigest.ComputeSha256(readbackStream.Value);
                if (!string.Equals(digest, plan.SourceArtifact.Sha256, StringComparison.OrdinalIgnoreCase))
                {
                    throw new InvalidOperationException("Fresh View rendering-resource digest failed: " + plan.TargetServerRelativeUrl);
                }
            }
            return true;
        }

        private static string TargetAuthoredReference(
            string sourceReference,
            ListViewRenderingResourceMaterializationPlan plan)
        {
            if (sourceReference.StartsWith("~sitecollection/", StringComparison.OrdinalIgnoreCase)
                || sourceReference.StartsWith("~site/", StringComparison.OrdinalIgnoreCase))
            {
                return sourceReference;
            }
            if (Uri.TryCreate(sourceReference, UriKind.Absolute, out var absolute))
            {
                return plan.TargetAbsoluteUrl + absolute.Query + absolute.Fragment;
            }
            if (sourceReference.StartsWith("/", StringComparison.Ordinal))
            {
                return plan.TargetServerRelativeUrl + Suffix(sourceReference);
            }
            return sourceReference;
        }

        private static string Suffix(string value)
        {
            var query = value.IndexOf('?');
            var fragment = value.IndexOf('#');
            var index = query < 0 ? fragment : fragment < 0 ? query : Math.Min(query, fragment);
            return index < 0 ? string.Empty : value.Substring(index);
        }

        private static Folder EnsureAssetFolder(ClientContext context, Web owner, string folderPath)
        {
            var separator = folderPath.IndexOf('/');
            var libraryPath = separator < 0 ? folderPath : folderPath.Substring(0, separator);
            var library = libraryPath.Equals("SiteAssets", StringComparison.OrdinalIgnoreCase)
                ? owner.Lists.EnsureSiteAssetsLibrary()
                : owner.GetListByUrl("Style Library");
            if (library == null)
            {
                throw new InvalidOperationException("The target Web does not expose the required '" + libraryPath + "' document library.");
            }
            context.Load(library.RootFolder, value => value.ServerRelativeUrl);
            context.ExecuteQueryRetry();
            var actualLibraryPath = RelativeToWeb(owner.ServerRelativeUrl, library.RootFolder.ServerRelativeUrl).Trim('/');
            if (!string.Equals(actualLibraryPath, libraryPath, StringComparison.OrdinalIgnoreCase))
            {
                throw new InvalidOperationException("The target library URL '" + actualLibraryPath + "' does not match sealed asset path '" + libraryPath + "'.");
            }
            return separator < 0 ? library.RootFolder : owner.EnsureFolderPath(folderPath);
        }

        private static Web ResolveOwner(Web pageWeb, Web rootWeb, string path)
        {
            var pagePath = pageWeb.ServerRelativeUrl == "/" ? string.Empty : pageWeb.ServerRelativeUrl.TrimEnd('/');
            var rootPath = rootWeb.ServerRelativeUrl == "/" ? string.Empty : rootWeb.ServerRelativeUrl.TrimEnd('/');
            if (!string.IsNullOrEmpty(pagePath) && path.StartsWith(pagePath + "/", StringComparison.OrdinalIgnoreCase))
            {
                return pageWeb;
            }
            if (string.IsNullOrEmpty(rootPath) || path.StartsWith(rootPath + "/", StringComparison.OrdinalIgnoreCase))
            {
                return rootWeb;
            }
            throw new InvalidOperationException("Target View rendering resource escapes the target Site Collection: " + path);
        }

        private static string RelativeToWeb(string webServerRelativeUrl, string path)
        {
            var prefix = webServerRelativeUrl == "/" ? string.Empty : webServerRelativeUrl.TrimEnd('/');
            if (!string.IsNullOrEmpty(prefix) && !path.StartsWith(prefix + "/", StringComparison.OrdinalIgnoreCase))
            {
                throw new InvalidOperationException("Target path '" + path + "' is outside Web '" + webServerRelativeUrl + "'.");
            }
            return path.Substring(prefix.Length).TrimStart('/');
        }

        private static void ValidateAssetFolder(string folderPath)
        {
            if (!(folderPath.Equals("SiteAssets", StringComparison.OrdinalIgnoreCase)
                || folderPath.StartsWith("SiteAssets/", StringComparison.OrdinalIgnoreCase)
                || folderPath.Equals("Style Library", StringComparison.OrdinalIgnoreCase)
                || folderPath.StartsWith("Style Library/", StringComparison.OrdinalIgnoreCase)))
            {
                throw new InvalidOperationException("Target View rendering resource must remain under SiteAssets or Style Library: " + folderPath);
            }
        }

        private static bool IsFileNotFound(ServerException exception)
        {
            return string.Equals(exception.ServerErrorTypeName, "System.IO.FileNotFoundException", StringComparison.Ordinal)
                || exception.ServerErrorCode == -2147024894;
        }
    }
}
