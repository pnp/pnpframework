using PnP.Framework.Migration.Packaging;
using System;

namespace PnP.Framework.Migration.Pages.Publishing.Layouts
{
    internal static class PublishingPageLayoutResourcePlanner
    {
        public static PublishingPageLayoutResourceMaterializationPlan Create(
            PublishingPageLayoutResourceSnapshot resource,
            Uri sourcePageWebUrl,
            Uri sourceSiteCollectionUrl,
            Uri targetPageWebUrl,
            Uri targetSiteCollectionUrl,
            bool allowExternalResourceReferences)
        {
            if (resource == null || resource.Reference == null)
            {
                throw new ArgumentNullException(nameof(resource));
            }

            if (resource.EvidenceState == PublishingPageLayoutResourceEvidenceState.TargetRuntime)
            {
                return new PublishingPageLayoutResourceMaterializationPlan
                {
                    SourceReference = resource.Reference.Value,
                    SourceEvidenceState = resource.EvidenceState,
                    Disposition = PublishingPageLayoutResourceMaterializationDisposition.TargetRuntime,
                    Reason = "The reference is a reviewed SharePoint target-runtime resource."
                };
            }

            Uri sourceUri;
            if (resource.EvidenceState != PublishingPageLayoutResourceEvidenceState.Readable
                || resource.Artifact == null
                || string.IsNullOrWhiteSpace(resource.ResolvedSourceUrl)
                || !Uri.TryCreate(resource.ResolvedSourceUrl, UriKind.Absolute, out sourceUri))
            {
                if (TryCreatePreservedExternal(resource, allowExternalResourceReferences, out var preserved))
                {
                    return preserved;
                }

                return Block(resource, "Exact source bytes and an absolute source URI are required before copying a Page Layout resource.");
            }

            Uri targetOwner;
            string relativePath;
            if (TryGetOwnedAssetRelativePath(sourcePageWebUrl, sourceUri, out relativePath))
            {
                targetOwner = targetPageWebUrl;
            }
            else if (TryGetOwnedAssetRelativePath(sourceSiteCollectionUrl, sourceUri, out relativePath))
            {
                targetOwner = targetSiteCollectionUrl;
            }
            else
            {
                if (TryCreatePreservedExternal(resource, allowExternalResourceReferences, out var preserved))
                {
                    return preserved;
                }

                return Block(resource, "Only source Page Web or site-collection SiteAssets and Style Library files have a reviewed target mapping.");
            }

            var targetPath = targetOwner.AbsolutePath.TrimEnd('/') + "/" + relativePath;
            var targetReference = new Uri(new Uri(targetOwner.GetLeftPart(UriPartial.Authority)), targetPath).AbsoluteUri;
            return new PublishingPageLayoutResourceMaterializationPlan
            {
                SourceReference = resource.Reference.Value,
                SourceUrl = sourceUri.AbsoluteUri,
                SourceEvidenceState = resource.EvidenceState,
                Disposition = PublishingPageLayoutResourceMaterializationDisposition.CreateOrReuseOwned,
                SourceArtifact = Copy(resource.Artifact),
                SourceContentBase64 = resource.ContentBase64,
                TargetServerRelativeUrl = Uri.UnescapeDataString(targetPath),
                TargetReference = targetReference,
                Reason = "Copy exact source bytes create-only to the corresponding target-owned asset path and rewrite the Page Layout reference."
            };
        }

        private static bool TryCreatePreservedExternal(
            PublishingPageLayoutResourceSnapshot resource,
            bool allowExternalResourceReferences,
            out PublishingPageLayoutResourceMaterializationPlan plan)
        {
            plan = null;
            if (!allowExternalResourceReferences
                || string.IsNullOrWhiteSpace(resource.Reference?.Value)
                || !Uri.TryCreate(resource.Reference.Value.Trim(), UriKind.Absolute, out var sourceUri)
                || !string.Equals(sourceUri.Scheme, Uri.UriSchemeHttps, StringComparison.OrdinalIgnoreCase))
            {
                return false;
            }

            plan = new PublishingPageLayoutResourceMaterializationPlan
            {
                SourceReference = resource.Reference.Value,
                SourceUrl = sourceUri.AbsoluteUri,
                SourceEvidenceState = resource.EvidenceState,
                Disposition = PublishingPageLayoutResourceMaterializationDisposition.PreserveExternal,
                TargetReference = resource.Reference.Value,
                Reason = "Retain the exact authored absolute HTTPS reference because external references are allowed. "
                    + $"The source payload remains '{resource.EvidenceState}', so the plan preserves the relationship without claiming or copying bytes."
            };
            return true;
        }

        private static PublishingPageLayoutResourceMaterializationPlan Block(
            PublishingPageLayoutResourceSnapshot resource,
            string reason)
        {
            return new PublishingPageLayoutResourceMaterializationPlan
            {
                SourceReference = resource.Reference?.Value,
                SourceUrl = resource.ResolvedSourceUrl,
                SourceEvidenceState = resource.EvidenceState,
                Disposition = PublishingPageLayoutResourceMaterializationDisposition.Block,
                SourceArtifact = resource.Artifact == null ? null : Copy(resource.Artifact),
                SourceContentBase64 = resource.ContentBase64,
                Reason = reason
            };
        }

        private static bool TryGetOwnedAssetRelativePath(Uri ownerWebUrl, Uri candidate, out string relativePath)
        {
            relativePath = null;
            if (!PublishingPageLayoutResourcePolicy.IsWebOwnedAsset(ownerWebUrl, candidate))
            {
                return false;
            }

            var ownerPath = Uri.UnescapeDataString(ownerWebUrl.AbsolutePath).TrimEnd('/');
            var sourcePath = Uri.UnescapeDataString(candidate.AbsolutePath);
            relativePath = sourcePath.Substring(ownerPath.Length).TrimStart('/');
            return true;
        }

        private static ArtifactReference Copy(ArtifactReference source)
        {
            return new ArtifactReference
            {
                Sha256 = source.Sha256,
                Length = source.Length,
                MediaType = source.MediaType,
                ContentEncoding = source.ContentEncoding,
                OriginalName = source.OriginalName,
                Availability = source.Availability,
                Lineage = source.Lineage
            };
        }
    }
}
