using Microsoft.SharePoint.Client;
using PnP.Framework.Migration.Packaging;
using System;
using System.IO;
using System.Linq;

namespace PnP.Framework.Migration.Pages.Publishing.Layouts
{
    internal static class PublishingPageLayoutResourceMaterializer
    {
        public static int Ensure(
            ClientContext context,
            PublishingPageLayoutMaterializationPlan plan,
            IMigrationArtifactStore artifactStore)
        {
            var resources = plan.ResourceMaterializations
                .Where(value => value.Disposition == PublishingPageLayoutResourceMaterializationDisposition.CreateOrReuseOwned)
                .GroupBy(value => value.TargetServerRelativeUrl, StringComparer.OrdinalIgnoreCase)
                .Select(group => group.First())
                .ToArray();
            var created = 0;
            foreach (var resource in resources)
            {
                if (resource.SourceArtifact == null || string.IsNullOrWhiteSpace(resource.TargetServerRelativeUrl))
                {
                    throw new InvalidOperationException($"Layout resource plan is incomplete: {resource.SourceReference}");
                }

                var bytes = MigrationArtifact.ReadAllBytes(resource.SourceArtifact, resource.SourceContentBase64, artifactStore);
                if (EnsureResource(context, resource, bytes))
                {
                    created++;
                }
            }

            return created;
        }

        private static bool EnsureResource(
            ClientContext context,
            PublishingPageLayoutResourceMaterializationPlan plan,
            byte[] bytes)
        {
            var rootWeb = context.Site.RootWeb;
            var pageWeb = context.Web;
            context.Load(rootWeb, value => value.ServerRelativeUrl, value => value.EffectiveBasePermissions);
            context.Load(pageWeb, value => value.ServerRelativeUrl, value => value.EffectiveBasePermissions);
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
                    string existingDigest;
                    using (stream.Value)
                    {
                        existingDigest = MigrationDigest.ComputeSha256(stream.Value);
                    }

                    if (!string.Equals(existingDigest, plan.SourceArtifact.Sha256, StringComparison.OrdinalIgnoreCase))
                    {
                        throw new InvalidOperationException($"Target layout resource collision: {plan.TargetServerRelativeUrl}");
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
                    throw new InvalidOperationException($"SharePoint did not persist layout resource '{plan.TargetServerRelativeUrl}'.");
                }
            }

            var readback = owner.GetFileByServerRelativePath(ResourcePath.FromDecodedUrl(plan.TargetServerRelativeUrl));
            var readbackStream = readback.OpenBinaryStream();
            context.Load(readback, value => value.Exists);
            context.ExecuteQueryRetry();
            if (!readback.Exists || readbackStream.Value == null)
            {
                throw new InvalidOperationException($"Fresh layout resource readback failed: {plan.TargetServerRelativeUrl}");
            }

            using (readbackStream.Value)
            {
                var digest = MigrationDigest.ComputeSha256(readbackStream.Value);
                if (!string.Equals(digest, plan.SourceArtifact.Sha256, StringComparison.OrdinalIgnoreCase))
                {
                    throw new InvalidOperationException($"Fresh layout resource readback digest failed: {plan.TargetServerRelativeUrl}");
                }
            }

            return true;
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
                throw new InvalidOperationException(
                    $"The target Web does not expose the required '{libraryPath}' document library.");
            }

            context.Load(library.RootFolder, value => value.ServerRelativeUrl);
            context.ExecuteQueryRetry();
            var actualLibraryPath = RelativeToWeb(owner.ServerRelativeUrl, library.RootFolder.ServerRelativeUrl).Trim('/');
            if (!string.Equals(actualLibraryPath, libraryPath, StringComparison.OrdinalIgnoreCase))
            {
                throw new InvalidOperationException(
                    $"The target library URL '{actualLibraryPath}' does not match sealed asset path '{libraryPath}'.");
            }

            return separator < 0 ? library.RootFolder : owner.EnsureFolderPath(folderPath);
        }

        private static Web ResolveOwner(Web pageWeb, Web rootWeb, string path)
        {
            var pagePath = pageWeb.ServerRelativeUrl.TrimEnd('/');
            var rootPath = rootWeb.ServerRelativeUrl == "/" ? string.Empty : rootWeb.ServerRelativeUrl.TrimEnd('/');
            if (!string.IsNullOrEmpty(pagePath) && path.StartsWith(pagePath + "/", StringComparison.OrdinalIgnoreCase))
            {
                return pageWeb;
            }

            if (string.IsNullOrEmpty(rootPath) || path.StartsWith(rootPath + "/", StringComparison.OrdinalIgnoreCase))
            {
                return rootWeb;
            }

            throw new InvalidOperationException($"Target layout resource escapes the target site collection: {path}");
        }

        private static string RelativeToWeb(string webServerRelativeUrl, string path)
        {
            var prefix = webServerRelativeUrl == "/" ? string.Empty : webServerRelativeUrl.TrimEnd('/');
            if (!string.IsNullOrEmpty(prefix) && !path.StartsWith(prefix + "/", StringComparison.OrdinalIgnoreCase))
            {
                throw new InvalidOperationException($"Target path '{path}' is outside Web '{webServerRelativeUrl}'.");
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
                throw new InvalidOperationException($"Target layout resource must remain under SiteAssets or Style Library: {folderPath}");
            }
        }

        private static bool IsFileNotFound(ServerException exception)
        {
            return string.Equals(exception.ServerErrorTypeName, "System.IO.FileNotFoundException", StringComparison.Ordinal)
                || exception.ServerErrorCode == -2147024894;
        }
    }
}
