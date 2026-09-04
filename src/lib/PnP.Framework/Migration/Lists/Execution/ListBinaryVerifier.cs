using Microsoft.SharePoint.Client;
using PnP.Framework.Migration.Lists.Capture;
using PnP.Framework.Migration.Lists.Items;
using PnP.Framework.Migration.Lists.Planning;
using PnP.Framework.Migration.Packaging;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace PnP.Framework.Migration.Lists.Execution
{
    internal static class ListBinaryVerifier
    {
        public static void VerifyDocument(
            ClientContext context,
            ListDependencySnapshot source,
            ListMaterializationPlan plan,
            ListItemSnapshot sourceItem,
            ListMaterializationReceipt receipt,
            ICollection<string> diagnostics)
        {
            if (sourceItem.Document == null)
            {
                return;
            }
            var path = ListDocumentMaterializer.MapPath(sourceItem.Document.ServerRelativeUrl, source.RootFolderServerRelativeUrl, plan.TargetRootFolderServerRelativeUrl);
            if (sourceItem.Document.Kind == ListDocumentObjectKind.Folder)
            {
                var folder = context.Web.GetFolderByServerRelativePath(ResourcePath.FromDecodedUrl(path));
                try
                {
                    context.Load(folder, value => value.Exists, value => value.ServerRelativeUrl);
                    context.ExecuteQueryRetry();
                    if (!folder.Exists)
                    {
                        diagnostics.Add("Target folder is missing: " + path + ".");
                    }
                    else
                    {
                        receipt.VerifiedDocumentCount++;
                    }
                }
                catch (Exception exception) when (exception is ServerException || exception is ClientRequestException)
                {
                    diagnostics.Add("Target folder could not be verified: " + path + ": " + exception.Message);
                }
                return;
            }
            var file = context.Web.GetFileByServerRelativePath(ResourcePath.FromDecodedUrl(path));
            if (VerifyFile(context, file, sourceItem.Document.Content == null ? null : sourceItem.Document.Content.Artifact, path, diagnostics))
            {
                receipt.VerifiedDocumentCount++;
            }
        }

        public static void VerifyAttachments(
            ClientContext context,
            ListItemSnapshot sourceItem,
            ListItem targetItem,
            ListMaterializationReceipt receipt,
            bool exactInventory,
            ICollection<string> diagnostics)
        {
            var actual = targetItem.AttachmentFiles.AsEnumerable().ToDictionary(value => value.FileName, StringComparer.OrdinalIgnoreCase);
            if (exactInventory && actual.Count != sourceItem.Attachments.Count)
            {
                diagnostics.Add("Target attachment count differs for source item " + sourceItem.SourceItemId + ".");
            }
            foreach (var attachment in sourceItem.Attachments)
            {
                Attachment target;
                if (!actual.TryGetValue(attachment.FileName, out target))
                {
                    diagnostics.Add("Target attachment is missing for source item " + sourceItem.SourceItemId + ": " + attachment.FileName + ".");
                    continue;
                }
                var file = context.Web.GetFileByServerRelativePath(ResourcePath.FromDecodedUrl(Uri.UnescapeDataString(target.ServerRelativeUrl)));
                if (VerifyFile(context, file, attachment.Content == null ? null : attachment.Content.Artifact, target.ServerRelativeUrl, diagnostics))
                {
                    receipt.VerifiedAttachmentCount++;
                }
            }
        }

        private static bool VerifyFile(
            ClientContext context,
            Microsoft.SharePoint.Client.File file,
            ArtifactReference expected,
            string path,
            ICollection<string> diagnostics)
        {
            if (expected == null)
            {
                diagnostics.Add("Expected binary artifact descriptor is missing: " + path + ".");
                return false;
            }
            try
            {
                var stream = file.OpenBinaryStream();
                context.ExecuteQueryRetry();
                using (stream.Value)
                using (var buffer = new MemoryStream())
                {
                    stream.Value.CopyTo(buffer);
                    var bytes = buffer.ToArray();
                    if (bytes.LongLength != expected.Length
                        || !string.Equals(MigrationDigest.ComputeSha256(bytes), expected.Sha256, StringComparison.OrdinalIgnoreCase))
                    {
                        diagnostics.Add("Target file bytes differ: " + path + ".");
                        return false;
                    }
                }
                return true;
            }
            catch (Exception exception) when (exception is ServerException || exception is ClientRequestException)
            {
                diagnostics.Add("Target file could not be verified: " + path + ": " + exception.Message);
                return false;
            }
        }
    }
}
