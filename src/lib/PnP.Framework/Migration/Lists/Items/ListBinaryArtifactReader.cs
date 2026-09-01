using Microsoft.SharePoint.Client;
using PnP.Framework.Migration.Evidence;
using PnP.Framework.Migration.Packaging;
using System;
using System.IO;

namespace PnP.Framework.Migration.Lists.Items
{
    internal static class ListBinaryArtifactReader
    {
        public static ListBinaryArtifactSnapshot Read(
            ClientContext context,
            Microsoft.SharePoint.Client.File file,
            long maximumBytes,
            IMigrationArtifactStore artifactStore,
            string mediaType,
            string originalName)
        {
            try
            {
                var streamResult = file.OpenBinaryStream();
                context.ExecuteQueryRetry();
                if (streamResult.Value == null)
                {
                    throw new FileNotFoundException("SharePoint returned no binary stream.");
                }

                byte[] bytes;
                using (streamResult.Value)
                using (var buffer = new MemoryStream())
                {
                    var block = new byte[81920];
                    int read;
                    while ((read = streamResult.Value.Read(block, 0, block.Length)) > 0)
                    {
                        buffer.Write(block, 0, read);
                        if (buffer.Length > maximumBytes)
                        {
                            throw new InvalidOperationException("The list binary artifact exceeds the configured maximum dependency size.");
                        }
                    }
                    bytes = buffer.ToArray();
                }

                ArtifactReference reference;
                string contentBase64 = null;
                if (artifactStore == null)
                {
                    reference = MigrationArtifact.Describe(bytes, mediaType, originalName);
                    contentBase64 = Convert.ToBase64String(bytes);
                }
                else
                {
                    using (var content = new MemoryStream(bytes, false))
                    {
                        reference = artifactStore.Put(content, mediaType, originalName);
                    }
                }
                return new ListBinaryArtifactSnapshot { Artifact = reference, ContentBase64 = contentBase64 };
            }
            catch (Exception exception) when (exception is ServerException || exception is IOException || exception is InvalidOperationException)
            {
                return new ListBinaryArtifactSnapshot
                {
                    Availability = EvidenceAvailability.Unavailable,
                    Diagnostics = { exception.Message }
                };
            }
        }

        public static string MediaType(string path)
        {
            switch (Path.GetExtension(path ?? string.Empty).ToLowerInvariant())
            {
                case ".css": return "text/css";
                case ".js": return "application/javascript";
                case ".json": return "application/json";
                case ".png": return "image/png";
                case ".jpg":
                case ".jpeg": return "image/jpeg";
                case ".gif": return "image/gif";
                case ".svg": return "image/svg+xml";
                case ".xsl":
                case ".xslt": return "application/xml";
                case ".pdf": return "application/pdf";
                default: return "application/octet-stream";
            }
        }
    }
}
