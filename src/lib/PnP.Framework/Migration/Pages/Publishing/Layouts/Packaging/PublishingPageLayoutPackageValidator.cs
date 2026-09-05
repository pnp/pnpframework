using PnP.Framework.Migration.Evidence;
using PnP.Framework.Migration.Diagnostics;
using PnP.Framework.Migration.Packaging;
using PnP.Framework.Migration.Schema.ContentTypes;
using PnP.Framework.Migration.Schema.ContentTypes.Packaging;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace PnP.Framework.Migration.Pages.Publishing.Layouts.Packaging
{
    internal static class PublishingPageLayoutPackageValidator
    {
        public static void ValidateSnapshot(
            PublishingPageLayoutSnapshot layout,
            IMigrationArtifactStore artifactStore)
        {
            if (layout == null
                || layout.Registrations == null
                || layout.Controls == null
                || layout.Zones == null
                || layout.ResourceReferences == null
                || layout.ResourceArtifacts == null
                || layout.Diagnostics == null)
            {
                throw new InvalidDataException("The Page Layout snapshot contains a null evidence collection.");
            }

            if (layout.EvidenceState == PublishingPageLayoutEvidenceState.Readable)
            {
                if (layout.Availability != EvidenceAvailability.Captured || layout.Bytes == null)
                {
                    throw new InvalidDataException("A readable Page Layout must contain captured byte evidence.");
                }

                MigrationArtifactContractValidator.Validate(layout.Bytes, layout.ContentBase64, artifactStore, "Page Layout");
            }
            else if (!string.IsNullOrWhiteSpace(layout.ContentBase64))
            {
                if (layout.Bytes == null)
                {
                    throw new InvalidDataException("A Page Layout inline payload requires an artifact reference.");
                }

                MigrationArtifactContractValidator.Validate(layout.Bytes, layout.ContentBase64, artifactStore, "Page Layout");
            }

            if (layout.ExternalToPageSiteCollection && string.IsNullOrWhiteSpace(layout.OwnerSiteCollectionUrl))
            {
                throw new InvalidDataException("An external Page Layout must identify its owner Site Collection.");
            }

            if (!string.IsNullOrWhiteSpace(layout.OwnerSiteCollectionUrl)
                && (!Uri.TryCreate(layout.OwnerSiteCollectionUrl, UriKind.Absolute, out var ownerSite)
                    || !string.Equals(ownerSite.Scheme, Uri.UriSchemeHttps, StringComparison.OrdinalIgnoreCase)))
            {
                throw new InvalidDataException("The Page Layout owner Site Collection URL must be absolute HTTPS.");
            }

            if (layout.EvidenceState == PublishingPageLayoutEvidenceState.AuthorizationBlocked)
            {
                LiteralHttpAuthorizationEvidence.Validate(layout.AuthorizationEvidence);
            }
            else if (layout.AuthorizationEvidence != null)
            {
                throw new InvalidDataException(
                    "Page Layout authorization evidence is only valid for the AuthorizationBlocked evidence state.");
            }

            var duplicateReference = layout.ResourceReferences
                .GroupBy(ResourceKey, StringComparer.OrdinalIgnoreCase)
                .FirstOrDefault(group => string.IsNullOrWhiteSpace(group.Key) || group.Count() > 1);
            if (duplicateReference != null)
            {
                throw new InvalidDataException("The Page Layout resource-reference inventory contains a missing or duplicate entry.");
            }

            if (layout.Registrations.Any(item => item == null)
                || layout.Controls.Any(item => item == null)
                || layout.Zones.Any(item => item == null))
            {
                throw new InvalidDataException("The Page Layout registration, control, or zone inventory contains a null entry.");
            }

            var duplicateResource = layout.ResourceArtifacts
                .GroupBy(item => item?.Reference == null ? null : ResourceKey(item.Reference), StringComparer.OrdinalIgnoreCase)
                .FirstOrDefault(group => string.IsNullOrWhiteSpace(group.Key) || group.Count() > 1);
            if (duplicateResource != null)
            {
                throw new InvalidDataException("The Page Layout resource evidence contains a missing or duplicate reference.");
            }

            var referenceKeys = new HashSet<string>(layout.ResourceReferences.Select(ResourceKey), StringComparer.OrdinalIgnoreCase);
            var artifactKeys = new HashSet<string>(layout.ResourceArtifacts.Select(item => ResourceKey(item.Reference)), StringComparer.OrdinalIgnoreCase);
            if (!referenceKeys.SetEquals(artifactKeys))
            {
                throw new InvalidDataException("The Page Layout must contain exactly one resource evidence record for every parsed reference.");
            }

            foreach (var resource in layout.ResourceArtifacts)
            {
                if (resource.Diagnostics == null || resource.Sources == null)
                {
                    throw new InvalidDataException($"The Page Layout resource '{resource.Reference.Value}' contains a null evidence collection.");
                }

                if (resource.EvidenceState == PublishingPageLayoutResourceEvidenceState.Readable)
                {
                    if (resource.Artifact == null)
                    {
                        throw new InvalidDataException($"Readable Page Layout resource '{resource.Reference.Value}' has no byte artifact.");
                    }

                    MigrationArtifactContractValidator.Validate(resource.Artifact, resource.ContentBase64, artifactStore,
                        $"Page Layout resource '{resource.Reference.Value}'");
                }
                else if (!string.IsNullOrWhiteSpace(resource.ContentBase64))
                {
                    if (resource.Artifact == null)
                    {
                        throw new InvalidDataException($"Page Layout resource '{resource.Reference.Value}' has inline bytes without an artifact reference.");
                    }

                    MigrationArtifactContractValidator.Validate(resource.Artifact, resource.ContentBase64, artifactStore,
                        $"Page Layout resource '{resource.Reference.Value}'");
                }
            }

            if (layout.AssociatedContentTypeSchema != null)
            {
                ContentTypeSchemaContractValidator.ValidateSnapshot(layout.AssociatedContentTypeSchema);
            }
        }

        public static void ValidatePlan(
            string pageLayoutName,
            bool isExecutable,
            PublishingPageLayoutMaterializationPlan layout,
            PublishingPageLayoutTargetProbe probe,
            PublishingPageLayoutTargetAdmission admission)
        {
            if (layout == null
                || admission == null
                || layout.RequiredFieldBindings == null
                || layout.RequiredRegistrations == null
                || layout.Zones == null
                || layout.ResourceReferences == null
                || layout.ResourceMaterializations == null
                || layout.ResourceRewrites == null
                || admission.Issues == null
                || admission.Warnings == null)
            {
                throw new InvalidDataException("The Page Layout plan or admission contains a null collection.");
            }

            if (!string.Equals(pageLayoutName, layout.TargetPageLayoutName, StringComparison.Ordinal))
            {
                throw new InvalidDataException("The page creation layout name does not match the sealed Page Layout materialization plan.");
            }

            if (isExecutable
                && (!admission.IsEligible
                    || admission.Disposition == PublishingPageLayoutMaterializationDisposition.Block
                    || probe == null))
            {
                var issueSummary = string.Join(" | ", (admission.Issues ?? new List<MigrationIssue>())
                    .Select(value => value.Code + ": " + value.Message));
                var probeDiagnostics = string.Join(" | ", probe?.Diagnostics ?? new List<string>());
                throw new InvalidDataException(
                    "An executable migration plan requires an eligible Page Layout admission and target probe. "
                    + "layoutDisposition=" + layout.Disposition
                    + "; admissionEligible=" + admission.IsEligible
                    + "; admissionDisposition=" + admission.Disposition
                    + "; probePresent=" + (probe != null)
                    + "; issues=" + issueSummary
                    + "; probeDiagnostics=" + probeDiagnostics);
            }

            if (probe != null
                && (probe.MissingFieldBindings == null
                    || probe.Resources == null
                    || probe.Diagnostics == null))
            {
                throw new InvalidDataException("The Page Layout target probe contains a null evidence collection.");
            }

            if (admission.Disposition != PublishingPageLayoutMaterializationDisposition.ReuseTargetStock
                && admission.IsEligible
                && (layout.ContentTypeSchema == null
                    || admission.ContentTypeSchema == null
                    || !admission.ContentTypeSchema.IsEligible))
            {
                throw new InvalidDataException("An eligible custom Page Layout requires an eligible associated content type schema admission.");
            }

            var duplicateResource = layout.ResourceMaterializations
                .Where(value => value != null && !string.IsNullOrWhiteSpace(value.TargetServerRelativeUrl))
                .GroupBy(value => value.SourceReference + "\u001f" + value.TargetServerRelativeUrl, StringComparer.OrdinalIgnoreCase)
                .FirstOrDefault(group => group.Count() > 1);
            if (layout.ResourceMaterializations.Any(value => value == null) || duplicateResource != null)
            {
                throw new InvalidDataException("The Page Layout materialization plan contains a null or duplicate resource action.");
            }

            if (layout.TargetBytes != null
                && layout.Disposition != PublishingPageLayoutMaterializationDisposition.ReuseTargetStock
                && layout.ResourceMaterializations.Count != layout.ResourceReferences.Count)
            {
                throw new InvalidDataException("A readable custom Page Layout plan must contain exactly one resource action for every captured reference.");
            }

            var conflictingTarget = layout.ResourceMaterializations
                .Where(value => value != null && !string.IsNullOrWhiteSpace(value.TargetServerRelativeUrl))
                .GroupBy(value => value.TargetServerRelativeUrl, StringComparer.OrdinalIgnoreCase)
                .FirstOrDefault(group => group
                    .Select(value => value.SourceArtifact?.Sha256)
                    .Distinct(StringComparer.OrdinalIgnoreCase)
                    .Count() > 1);
            if (conflictingTarget != null)
            {
                throw new InvalidDataException($"Multiple Page Layout resources map different bytes to target path '{conflictingTarget.Key}'.");
            }

            if (layout.ContentTypeSchema != null)
            {
                ContentTypeSchemaContractValidator.ValidatePlan(layout.ContentTypeSchema);
            }
        }

        private static string ResourceKey(PublishingPageLayoutResourceReference reference)
        {
            return reference == null
                ? null
                : (reference.Attribute ?? string.Empty) + "\u001f" + (reference.Value ?? string.Empty);
        }
    }
}
