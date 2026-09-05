using Microsoft.SharePoint.Client;
using PnP.Framework.Migration.Lists.Items;
using PnP.Framework.Migration.Packaging;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace PnP.Framework.Migration.Lists.Execution
{
    internal static class ListAttachmentMaterializer
    {
        public static void Ensure(
            ClientContext context,
            ListItem item,
            IEnumerable<ListAttachmentSnapshot> attachments,
            IMigrationArtifactStore artifactStore)
        {
            var values = attachments.ToArray();
            if (values.Length == 0)
            {
                return;
            }
            context.Load(item.AttachmentFiles, files => files.Include(value => value.FileName, value => value.ServerRelativeUrl));
            context.ExecuteQueryRetry();
            var existing = item.AttachmentFiles.ToDictionary(value => value.FileName, StringComparer.OrdinalIgnoreCase);
            foreach (var attachment in values)
            {
                Attachment existingAttachment;
                if (existing.TryGetValue(attachment.FileName, out existingAttachment))
                {
                    Microsoft.SharePoint.Client.File existingFile;
                    if (!ListBinaryMaterializer.TryGetFile(context, existingAttachment.ServerRelativeUrl, out existingFile))
                    {
                        throw new InvalidDataException("Existing target attachment could not be opened for verification: " + existingAttachment.ServerRelativeUrl);
                    }
                    ListBinaryMaterializer.VerifyExistingFile(context, existingFile, attachment.Content.Artifact);
                    continue;
                }
                var bytes = MigrationArtifact.ReadAllBytes(attachment.Content.Artifact, attachment.Content.ContentBase64, artifactStore);
                using (var stream = new MemoryStream(bytes, false))
                {
                    item.AttachmentFiles.Add(new AttachmentCreationInformation { FileName = attachment.FileName, ContentStream = stream });
                    context.ExecuteQueryRetry();
                }
            }
        }
    }
}
