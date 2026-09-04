using Microsoft.SharePoint.Client;
using PnP.Framework.Migration.Execution;
using PnP.Framework.Migration.Pages.Publishing.Packaging;
using System;
using System.IO;

namespace PnP.Framework.Migration.Pages.Publishing.Execution
{
    internal sealed class PublishingPageTargetLocation
    {
        public List PagesLibrary { get; set; }

        public Folder TargetFolder { get; set; }

        public string FileName { get; set; }
    }

    /// <summary>
    /// Resolves the approved Pages library and creates only the missing folder
    /// nodes needed by the exact target Page path. It never changes the planned
    /// library, folder hierarchy, or filename.
    /// </summary>
    internal static class PublishingPageTargetLocationMaterializer
    {
        public static PublishingPageTargetLocation Ensure(
            ClientContext context,
            PublishingPageMigrationPackage package,
            MigrationExecutionRecorder recorder)
        {
            if (context == null)
            {
                throw new ArgumentNullException(nameof(context));
            }
            if (package?.Plan?.TargetProbe == null)
            {
                throw new InvalidDataException("The sealed target Pages-library probe is required.");
            }
            if (recorder == null)
            {
                throw new ArgumentNullException(nameof(recorder));
            }

            var plannedRoot = Normalize(package.Plan.TargetProbe.PagesLibraryServerRelativeUrl);
            var targetPath = Normalize(package.Plan.TargetPageServerRelativeUrl);
            if (string.IsNullOrWhiteSpace(plannedRoot) || string.IsNullOrWhiteSpace(targetPath))
            {
                throw new InvalidDataException("The approved Pages-library and target Page paths are required.");
            }

            var web = context.Web;
            var pages = web.GetList(plannedRoot);
            context.Load(web, value => value.ServerRelativeUrl);
            context.Load(pages, value => value.EnableModeration, value => value.ForceCheckout);
            context.Load(pages.RootFolder, value => value.ServerRelativeUrl);
            context.ExecuteQueryRetry();

            var actualRoot = Normalize(pages.RootFolder.ServerRelativeUrl);
            if (!string.Equals(plannedRoot, actualRoot, StringComparison.OrdinalIgnoreCase))
            {
                throw new InvalidDataException(
                    $"The target Pages-library root changed after approval: expected '{plannedRoot}', observed '{actualRoot}'.");
            }

            var targetDirectory = Normalize(PagePath.GetDirectoryName(targetPath));
            if (!PagePath.IsWithin(targetDirectory, actualRoot))
            {
                throw new InvalidDataException(
                    $"The approved Page path '{targetPath}' is outside the approved Pages library '{actualRoot}'.");
            }

            Folder targetFolder = null;
            if (string.Equals(targetDirectory, actualRoot, StringComparison.OrdinalIgnoreCase))
            {
                recorder.RecordAlreadySatisfied(
                    "page.folder",
                    "The approved Page path is at the Pages-library root; no folder transaction is required.");
            }
            else if (TryGetFolder(context, targetDirectory, out targetFolder))
            {
                recorder.RecordAlreadySatisfied(
                    "page.folder",
                    $"Reuse the existing exact target Page folder '{targetDirectory}'.");
            }
            else
            {
                var webRelativeDirectory = GetWebRelativeDirectory(web.ServerRelativeUrl, targetDirectory);
                targetFolder = recorder.Execute(
                    "page.folder",
                    $"Create the missing exact target Page folder hierarchy '{targetDirectory}'.",
                    () => web.EnsureFolderPath(webRelativeDirectory, value => value.ServerRelativeUrl),
                    value => MutationOutcome.Applied,
                    value => $"Created the missing Page folder hierarchy '{targetDirectory}'.");
            }

            if (targetFolder != null)
            {
                targetFolder.EnsureProperty(value => value.ServerRelativeUrl);
                if (!string.Equals(
                        Normalize(targetFolder.ServerRelativeUrl),
                        targetDirectory,
                        StringComparison.OrdinalIgnoreCase))
                {
                    throw new InvalidOperationException(
                        $"SharePoint resolved Page folder '{targetFolder.ServerRelativeUrl}' instead of the approved path '{targetDirectory}'.");
                }
            }

            return new PublishingPageTargetLocation
            {
                PagesLibrary = pages,
                TargetFolder = targetFolder,
                FileName = PagePath.GetFileName(targetPath)
            };
        }

        internal static string GetWebRelativeDirectory(string webServerRelativeUrl, string directory)
        {
            var webPath = Normalize(webServerRelativeUrl);
            var targetDirectory = Normalize(directory);
            if (!PagePath.IsWithin(targetDirectory, webPath))
            {
                throw new InvalidDataException(
                    $"Target directory '{targetDirectory}' is outside target Web '{webPath}'.");
            }
            if (string.Equals(webPath, "/", StringComparison.Ordinal))
            {
                return targetDirectory.TrimStart('/');
            }
            return targetDirectory.Substring(webPath.Length).TrimStart('/');
        }

        private static bool TryGetFolder(ClientContext context, string serverRelativeUrl, out Folder folder)
        {
            folder = context.Web.GetFolderByServerRelativePath(ResourcePath.FromDecodedUrl(serverRelativeUrl));
            context.Load(folder, value => value.ServerRelativeUrl);
            try
            {
                context.ExecuteQueryRetry();
                return true;
            }
            catch (ServerException exception) when (IsMissing(exception))
            {
                folder = null;
                return false;
            }
        }

        private static bool IsMissing(ServerException exception)
        {
            return string.Equals(
                    exception.ServerErrorTypeName,
                    "System.IO.FileNotFoundException",
                    StringComparison.Ordinal)
                || exception.ServerErrorCode == -2147024894;
        }

        private static string Normalize(string value)
        {
            if (string.IsNullOrWhiteSpace(value))
            {
                return value;
            }
            var result = Uri.UnescapeDataString(value).Replace('\\', '/').Trim();
            if (!result.StartsWith("/", StringComparison.Ordinal))
            {
                result = "/" + result;
            }
            return result.Length == 1 ? result : result.TrimEnd('/');
        }
    }
}
