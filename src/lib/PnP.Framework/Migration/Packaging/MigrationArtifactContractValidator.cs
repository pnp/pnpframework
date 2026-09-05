using System;
using System.IO;

namespace PnP.Framework.Migration.Packaging
{
    internal static class MigrationArtifactContractValidator
    {
        public static void Validate(
            ArtifactReference artifact,
            string contentBase64,
            IMigrationArtifactStore artifactStore,
            string description)
        {
            if (artifact == null
                || string.IsNullOrWhiteSpace(artifact.Sha256)
                || artifact.Sha256.Length != 64
                || artifact.Length < 0)
            {
                throw new InvalidDataException($"{description} artifact metadata is incomplete.");
            }

            if (!string.IsNullOrWhiteSpace(contentBase64))
            {
                byte[] bytes;
                try
                {
                    bytes = Convert.FromBase64String(contentBase64);
                }
                catch (FormatException exception)
                {
                    throw new InvalidDataException($"{description} inline payload is not valid Base64.", exception);
                }

                if (bytes.LongLength != artifact.Length
                    || !string.Equals(MigrationDigest.ComputeSha256(bytes), artifact.Sha256, StringComparison.OrdinalIgnoreCase))
                {
                    throw new InvalidDataException($"{description} inline payload length or digest does not match its artifact reference.");
                }

                return;
            }

            if (artifactStore == null)
            {
                return;
            }

            if (!artifactStore.Contains(artifact.Sha256))
            {
                throw new InvalidDataException($"{description} artifact '{artifact.Sha256}' is absent from the supplied artifact store.");
            }

            using (var content = artifactStore.OpenRead(artifact.Sha256))
            {
                var length = content.CanSeek ? content.Length - content.Position : (long?)null;
                var digest = MigrationDigest.ComputeSha256(content);
                if ((length.HasValue && length.Value != artifact.Length)
                    || !string.Equals(digest, artifact.Sha256, StringComparison.OrdinalIgnoreCase))
                {
                    throw new InvalidDataException($"{description} artifact-store payload differs from its sealed reference.");
                }
            }
        }
    }
}
