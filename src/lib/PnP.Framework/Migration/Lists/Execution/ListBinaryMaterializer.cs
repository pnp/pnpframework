using Microsoft.SharePoint.Client;
using PnP.Framework.Migration.Packaging;
using System;
using System.IO;

namespace PnP.Framework.Migration.Lists.Execution
{
    internal static class ListBinaryMaterializer
    {
        public static bool TryGetFile(ClientContext context, string path, out Microsoft.SharePoint.Client.File file)
        {
            file = context.Web.GetFileByServerRelativePath(ResourcePath.FromDecodedUrl(path));
            try
            {
                context.Load(file, value => value.Exists, value => value.Length, value => value.ServerRelativeUrl);
                context.ExecuteQueryRetry();
                return file.Exists;
            }
            catch (Exception exception) when (exception is ServerException || exception is ClientRequestException)
            {
                return false;
            }
        }

        public static void VerifyExistingFile(ClientContext context, Microsoft.SharePoint.Client.File file, ArtifactReference expected)
        {
            var stream = file.OpenBinaryStream();
            context.ExecuteQueryRetry();
            using (stream.Value)
            using (var buffer = new MemoryStream())
            {
                stream.Value.CopyTo(buffer);
                var bytes = buffer.ToArray();
                if (bytes.LongLength != expected.Length || !string.Equals(MigrationDigest.ComputeSha256(bytes), expected.Sha256, StringComparison.OrdinalIgnoreCase))
                {
                    throw new InvalidDataException("Existing target file bytes differ from the sealed source artifact: " + file.ServerRelativeUrl);
                }
            }
        }
    }
}
