using Microsoft.SharePoint.Client;
using PnP.Framework.Migration.Evidence;
using PnP.Framework.Migration.Packaging;
using System;
using System.Collections.Generic;
using System.IO;

namespace PnP.Framework.Migration.Pages.Publishing.Layouts
{
    internal static class PublishingPageLayoutResourceSnapshotReader
    {
        public static PublishingPageLayoutResourceSnapshot Read(
            ClientContext context,
            Uri sourceWebUrl,
            Uri sourceSiteCollectionUrl,
            PublishingPageLayoutResourceReference reference,
            IMigrationArtifactStore artifactStore)
        {
            if (PublishingPageLayoutResourcePolicy.IsTargetRuntimeResource(reference.Value))
            {
                return Result(reference, PublishingPageLayoutResourceEvidenceState.TargetRuntime, null, null, null);
            }

            var sourceUri = PublishingPageLayoutResourcePolicy.ResolveSourceUri(sourceWebUrl, sourceSiteCollectionUrl, reference.Value);
            if (sourceUri == null)
            {
                return Result(reference, PublishingPageLayoutResourceEvidenceState.Unsupported, null, null,
                    "The layout resource reference could not be resolved to an HTTPS source URI.");
            }

            if (!PublishingPageLayoutResourcePolicy.IsWebOwnedAsset(sourceWebUrl, sourceUri)
                && !PublishingPageLayoutResourcePolicy.IsWebOwnedAsset(sourceSiteCollectionUrl, sourceUri))
            {
                return Result(reference, PublishingPageLayoutResourceEvidenceState.Unsupported, sourceUri.AbsoluteUri, null,
                    "Only source Web or site-collection SiteAssets and Style Library resources have a reviewed copy path.");
            }

            try
            {
                var path = Uri.UnescapeDataString(sourceUri.AbsolutePath);
                var ownerWeb = PublishingPageLayoutResourcePolicy.IsWebOwnedAsset(sourceWebUrl, sourceUri)
                    ? context.Web
                    : context.Site.RootWeb;
                var file = ownerWeb.GetFileByServerRelativePath(ResourcePath.FromDecodedUrl(path));
                var stream = file.OpenBinaryStream();
                context.Load(file, value => value.Exists, value => value.Name, value => value.Length);
                context.ExecuteQueryRetry();
                if (!file.Exists || stream.Value == null)
                {
                    return Result(reference, PublishingPageLayoutResourceEvidenceState.Missing, sourceUri.AbsoluteUri, null,
                        "The layout resource file was not found or returned no binary stream.");
                }

                byte[] bytes;
                using (stream.Value)
                using (var buffer = new MemoryStream())
                {
                    stream.Value.CopyTo(buffer);
                    bytes = buffer.ToArray();
                }

                var artifact = artifactStore == null
                    ? MigrationArtifact.Describe(bytes, null, file.Name)
                    : Put(artifactStore, bytes, file.Name);
                return new PublishingPageLayoutResourceSnapshot
                {
                    Reference = reference,
                    EvidenceState = PublishingPageLayoutResourceEvidenceState.Readable,
                    ResolvedSourceUrl = sourceUri.AbsoluteUri,
                    Artifact = artifact,
                    ContentBase64 = artifactStore == null ? Convert.ToBase64String(bytes) : null,
                    Diagnostics = new List<string>()
                };
            }
            catch (ServerException exception)
            {
                var accessDenied = exception.ServerErrorCode == -2147024891
                    || exception.Message.IndexOf("Access denied", StringComparison.OrdinalIgnoreCase) >= 0;
                var missing = exception.ServerErrorCode == -2147024894
                    || string.Equals(exception.ServerErrorTypeName, "System.IO.FileNotFoundException", StringComparison.Ordinal);
                return Result(
                    reference,
                    accessDenied ? PublishingPageLayoutResourceEvidenceState.AccessDenied
                        : missing ? PublishingPageLayoutResourceEvidenceState.Missing
                        : PublishingPageLayoutResourceEvidenceState.Failed,
                    sourceUri.AbsoluteUri,
                    null,
                    exception.Message);
            }
        }

        private static ArtifactReference Put(IMigrationArtifactStore store, byte[] bytes, string name)
        {
            using (var content = new MemoryStream(bytes, false))
            {
                return store.Put(content, null, name);
            }
        }

        private static PublishingPageLayoutResourceSnapshot Result(
            PublishingPageLayoutResourceReference reference,
            PublishingPageLayoutResourceEvidenceState state,
            string sourceUrl,
            ArtifactReference artifact,
            string diagnostic)
        {
            return new PublishingPageLayoutResourceSnapshot
            {
                Reference = reference,
                EvidenceState = state,
                ResolvedSourceUrl = sourceUrl,
                Artifact = artifact,
                Diagnostics = string.IsNullOrWhiteSpace(diagnostic)
                    ? new List<string>()
                    : new List<string> { diagnostic },
                Sources = new List<EvidenceSource>()
            };
        }
    }
}
