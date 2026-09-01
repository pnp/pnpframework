using System;
using System.IO;

namespace PnP.Framework.Migration.Packaging
{
    public static class MigrationArtifact
    {
        public static ArtifactReference Describe(byte[] content, string mediaType = null, string originalName = null)
        {
            if (content == null)
            {
                throw new ArgumentNullException(nameof(content));
            }

            return new ArtifactReference
            {
                Sha256 = MigrationDigest.ComputeSha256(content),
                Length = content.LongLength,
                MediaType = mediaType,
                OriginalName = originalName
            };
        }

        public static byte[] ReadAllBytes(
            ArtifactReference reference,
            string contentBase64,
            IMigrationArtifactStore artifactStore = null)
        {
            if (reference == null)
            {
                throw new ArgumentNullException(nameof(reference));
            }

            byte[] content;
            if (!string.IsNullOrWhiteSpace(contentBase64))
            {
                try
                {
                    content = Convert.FromBase64String(contentBase64);
                }
                catch (FormatException exception)
                {
                    throw new InvalidDataException("The inline artifact payload is not valid Base64.", exception);
                }
            }
            else
            {
                if (artifactStore == null || !artifactStore.Contains(reference.Sha256))
                {
                    throw new InvalidDataException($"Artifact '{reference.Sha256}' is not available inline or in the supplied artifact store.");
                }

                using (var source = artifactStore.OpenRead(reference.Sha256))
                using (var buffer = new MemoryStream())
                {
                    source.CopyTo(buffer);
                    content = buffer.ToArray();
                }
            }

            if (content.LongLength != reference.Length
                || !string.Equals(MigrationDigest.ComputeSha256(content), reference.Sha256, StringComparison.OrdinalIgnoreCase))
            {
                throw new InvalidDataException($"Artifact '{reference.Sha256}' payload differs from its sealed length or digest.");
            }

            return content;
        }
    }
}
