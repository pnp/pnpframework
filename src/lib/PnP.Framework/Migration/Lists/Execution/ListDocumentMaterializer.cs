using Microsoft.SharePoint.Client;
using PnP.Framework.Migration.Lists.Capture;
using PnP.Framework.Migration.Lists.Items;
using PnP.Framework.Migration.Lists.Planning;
using PnP.Framework.Migration.Packaging;
using System;
using System.IO;

namespace PnP.Framework.Migration.Lists.Execution
{
    internal static class ListDocumentMaterializer
    {
        public static ListItem CreateItem(
            ClientContext context,
            List targetList,
            ListDependencySnapshot source,
            ListMaterializationPlan plan,
            ListItemSnapshot sourceItem,
            IMigrationArtifactStore artifactStore)
        {
            if (sourceItem.Document == null)
            {
                var item = targetList.AddItem(new ListItemCreationInformation());
                item[ListItemMaterializer.OriginalItemIdFieldName] = sourceItem.SourceItemId;
                item.Update();
                context.ExecuteQueryRetry();
                context.Load(item, value => value.Id);
                context.ExecuteQueryRetry();
                return item;
            }

            var targetPath = MapPath(sourceItem.Document.ServerRelativeUrl, source.RootFolderServerRelativeUrl, plan.TargetRootFolderServerRelativeUrl);
            if (sourceItem.Document.Kind == ListDocumentObjectKind.Folder)
            {
                var relative = targetPath.Substring(plan.TargetRootFolderServerRelativeUrl.TrimEnd('/').Length).Trim('/');
                var folder = context.Web.EnsureFolder(targetList.RootFolder, relative, value => value.ServerRelativeUrl, value => value.ListItemAllFields);
                var item = folder.ListItemAllFields;
                context.Load(item, value => value.Id);
                context.ExecuteQueryRetry();
                item[ListItemMaterializer.OriginalItemIdFieldName] = sourceItem.SourceItemId;
                item.Update();
                context.ExecuteQueryRetry();
                return item;
            }

            if (ListMigrationPlanFactory.IsRightsManagedEnvelope(sourceItem.Document.Content))
            {
                throw new InvalidOperationException(
                    "Rights-managed document replay requires an approved cross-site envelope materializer and semantic verifier: "
                    + sourceItem.Document.ServerRelativeUrl);
            }
            var bytes = MigrationArtifact.ReadAllBytes(sourceItem.Document.Content.Artifact, sourceItem.Document.Content.ContentBase64, artifactStore);
            var directory = targetPath.Substring(0, targetPath.LastIndexOf('/'));
            var relativeDirectory = directory.Substring(plan.TargetRootFolderServerRelativeUrl.TrimEnd('/').Length).Trim('/');
            var targetFolder = string.IsNullOrEmpty(relativeDirectory)
                ? targetList.RootFolder
                : context.Web.EnsureFolder(targetList.RootFolder, relativeDirectory, value => value.ServerRelativeUrl);
            Microsoft.SharePoint.Client.File file;
            if (ListBinaryMaterializer.TryGetFile(context, targetPath, out file))
            {
                ListBinaryMaterializer.VerifyExistingFile(context, file, sourceItem.Document.Content.Artifact);
            }
            else
            {
                using (var stream = new MemoryStream(bytes, false))
                {
                    file = targetFolder.UploadFile(sourceItem.Document.Name, stream, false);
                }
            }
            var listItem = file.ListItemAllFields;
            context.Load(listItem, value => value.Id);
            context.ExecuteQueryRetry();
            listItem[ListItemMaterializer.OriginalItemIdFieldName] = sourceItem.SourceItemId;
            listItem.Update();
            context.ExecuteQueryRetry();
            return listItem;
        }

        public static string MapPath(string sourcePath, string sourceRoot, string targetRoot)
        {
            if (!sourcePath.StartsWith(sourceRoot.TrimEnd('/') + "/", StringComparison.OrdinalIgnoreCase))
            {
                throw new InvalidDataException("Document path is outside its captured List root: " + sourcePath);
            }
            return targetRoot.TrimEnd('/') + sourcePath.Substring(sourceRoot.TrimEnd('/').Length);
        }
    }
}
