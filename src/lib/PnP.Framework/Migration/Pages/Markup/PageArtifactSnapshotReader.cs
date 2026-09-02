using Microsoft.SharePoint.Client;
using PnP.Framework.Migration.Evidence;
using PnP.Framework.Migration.Packaging;
using System;
using System.Collections.Generic;
using System.IO;

namespace PnP.Framework.Migration.Pages.Markup
{
    internal static class PageArtifactSnapshotReader
    {
        public static PageArtifactSnapshot Read(
            ClientContext context,
            Microsoft.SharePoint.Client.File file,
            IMigrationArtifactStore artifactStore,
            ICollection<string> blockers)
        {
            try
            {
                var stream = file.OpenBinaryStream();
                context.ExecuteQueryRetry();
                if (stream.Value == null)
                {
                    return Unavailable(file, "The source page ASPX bytes are unavailable.", blockers);
                }

                byte[] bytes;
                using (stream.Value)
                using (var buffer = new MemoryStream())
                {
                    stream.Value.CopyTo(buffer);
                    bytes = buffer.ToArray();
                }

                var artifact = artifactStore == null
                    ? MigrationArtifact.Describe(bytes, "application/vnd.ms-aspx", file.Name)
                    : Put(artifactStore, bytes, file.Name);
                return new PageArtifactSnapshot
                {
                    FileUniqueId = file.UniqueId,
                    ServerRelativeUrl = file.ServerRelativeUrl,
                    Bytes = artifact,
                    ContentBase64 = artifactStore == null ? Convert.ToBase64String(bytes) : null,
                    PageDirective = PageDirectiveParser.Parse(PageMarkupEncoding.Decode(bytes)),
                    Availability = EvidenceAvailability.Captured
                };
            }
            catch (ServerException exception)
            {
                return Unavailable(file, exception.Message, blockers);
            }
            catch (IOException exception)
            {
                return Unavailable(file, exception.Message, blockers);
            }
        }

        private static ArtifactReference Put(IMigrationArtifactStore store, byte[] bytes, string name)
        {
            using (var content = new MemoryStream(bytes, false))
            {
                return store.Put(content, "application/vnd.ms-aspx", name);
            }
        }

        private static PageArtifactSnapshot Unavailable(
            Microsoft.SharePoint.Client.File file,
            string diagnostic,
            ICollection<string> blockers)
        {
            blockers?.Add("Source page artifact capture failed: " + diagnostic);
            return new PageArtifactSnapshot
            {
                FileUniqueId = file.UniqueId,
                ServerRelativeUrl = file.ServerRelativeUrl,
                Availability = EvidenceAvailability.Unavailable,
                Diagnostics = new List<string> { diagnostic }
            };
        }
    }
}
