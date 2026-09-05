using PnP.Framework.Migration.Evidence;
using PnP.Framework.Migration.Lists.Capture;
using PnP.Framework.Migration.Lists.Items;
using PnP.Framework.Migration.Lists.Items.Protection;
using PnP.Framework.Migration.Lists.Planning;
using PnP.Framework.Migration.Packaging;
using PnP.Framework.Migration.Pages.ClassicWebParts;
using PnP.Framework.Migration.Pages.ClassicWebParts.Bindings;
using PnP.Framework.Migration.Topology;
using PnP.Framework.Migration.Schema.ContentTypes;
using PnP.Framework.Migration.Schema.ContentTypes.Packaging;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace PnP.Framework.Migration.Lists.Packaging
{
    internal static class ListDependencyPackageValidator
    {
        public static void Validate(
            IEnumerable<ClassicWebPartSnapshot> webParts,
            IEnumerable<ClassicListWebPartBindingSnapshot> bindings,
            IEnumerable<ListDependencySnapshot> dependencies,
            IEnumerable<ListLookupDependency> lookupDependencies,
            SourceSiteCollectionSnapshot topology,
            IMigrationArtifactStore artifactStore)
        {
            if (bindings == null || dependencies == null || lookupDependencies == null)
            {
                throw new InvalidDataException("The list dependency snapshot contains a null inventory collection.");
            }

            var partById = (webParts ?? Enumerable.Empty<ClassicWebPartSnapshot>()).ToDictionary(value => value.Id);
            var bindingValues = bindings.ToArray();
            var duplicateBinding = bindingValues.GroupBy(value => value == null ? Guid.Empty : value.SourceWebPartId).FirstOrDefault(group => group.Key == Guid.Empty || group.Count() > 1);
            if (duplicateBinding != null)
            {
                throw new InvalidDataException("The list Web Part binding inventory contains a missing or duplicate Web Part ID.");
            }

            foreach (var binding in bindingValues)
            {
                ClassicWebPartSnapshot webPart;
                if (!partById.TryGetValue(binding.SourceWebPartId, out webPart)
                    || binding.SourceListId == Guid.Empty || binding.SourceListWebId == Guid.Empty
                    || !string.Equals(webPart.ExportSha256, binding.SourceExportSha256, StringComparison.OrdinalIgnoreCase)
                    || !string.Equals(webPart.ExportXml, binding.SourceExportXml, StringComparison.Ordinal))
                {
                    throw new InvalidDataException("A list Web Part binding does not reference exact captured Web Part evidence: " + binding.SourceWebPartId.ToString("D"));
                }
            }

            var dependencyValues = dependencies.ToArray();
            var dependencyByList = dependencyValues.GroupBy(value => value == null ? Guid.Empty : value.SourceListId).ToDictionary(group => group.Key, group => group.ToArray());
            if (dependencyByList.Any(value => value.Key == Guid.Empty || value.Value.Length != 1))
            {
                throw new InvalidDataException("The list dependency inventory contains a missing or duplicate source List ID.");
            }
            foreach (var binding in bindingValues)
            {
                ListDependencySnapshot[] matches;
                if (!dependencyByList.TryGetValue(binding.SourceListId, out matches) || matches.Length != 1 || matches[0].SourceWebId != binding.SourceListWebId)
                {
                    throw new InvalidDataException("A list Web Part binding has no exact captured List dependency: " + binding.SourceListId.ToString("D"));
                }
            }

            foreach (var dependency in dependencyValues)
            {
                ValidateDependency(dependency, artifactStore);
            }
            foreach (var edge in lookupDependencies)
            {
                if (edge == null || edge.SourceListId == Guid.Empty || edge.LookupListId == Guid.Empty
                    || !dependencyByList.ContainsKey(edge.SourceListId) || !dependencyByList.ContainsKey(edge.LookupListId))
                {
                    throw new InvalidDataException("The lookup dependency graph references an uncaptured List.");
                }
                var owner = dependencyByList[edge.SourceListId][0];
                if (!owner.Fields.Any(value => value.Id == edge.FieldId && value.SourceLookupListId == edge.LookupListId))
                {
                    throw new InvalidDataException("The lookup dependency graph does not match captured field schema: " + edge.FieldId.ToString("D"));
                }
            }
            var order = ListLookupDependencyGraph.Order(dependencyValues.Select(value => value.SourceListId), lookupDependencies);
            if (!order.IsExecutable)
            {
                throw new InvalidDataException(order.Issues.First().Message);
            }

            if (bindingValues.Length > 0 && topology != null)
            {
                if (topology.Webs == null)
                {
                    throw new InvalidDataException("A captured source topology closure has a null Web inventory.");
                }
                var capturedWebIds = new HashSet<Guid>(topology.Webs.Select(value => value.WebId));
                if (dependencyValues.Any(value => value.SourceSiteId != topology.SiteId || !capturedWebIds.Contains(value.SourceWebId)))
                {
                    throw new InvalidDataException("A captured List dependency is outside the sealed source topology closure.");
                }
            }
        }

        private static void ValidateDependency(ListDependencySnapshot dependency, IMigrationArtifactStore artifactStore)
        {
            if (dependency == null || dependency.SourceSiteId == Guid.Empty || dependency.SourceWebId == Guid.Empty || dependency.SourceListId == Guid.Empty
                || string.IsNullOrWhiteSpace(dependency.SourceWebUrl) || string.IsNullOrWhiteSpace(dependency.RootFolderServerRelativeUrl)
                || dependency.Fields == null || dependency.ContentTypes == null || dependency.UniqueContentTypeOrder == null || dependency.SiteContentTypes == null
                || dependency.Views == null || dependency.ViewRenderingResources == null
                || dependency.Items == null || dependency.Diagnostics == null)
            {
                throw new InvalidDataException("A List dependency snapshot is missing identity, metadata, or an inventory collection.");
            }
            ValidateInformationRightsManagement(dependency);
            var duplicateField = dependency.Fields.GroupBy(value => value == null ? Guid.Empty : value.Id).FirstOrDefault(group => group.Key == Guid.Empty || group.Count() > 1);
            if (duplicateField != null || dependency.Fields.Any(value => string.IsNullOrWhiteSpace(value.InternalName)
                || !string.Equals(MigrationDigest.ComputeSha256(value.SchemaXml ?? string.Empty), value.SchemaXmlSha256, StringComparison.OrdinalIgnoreCase)))
            {
                throw new InvalidDataException("List '" + dependency.Title + "' contains missing, duplicate, or mutated field schema evidence.");
            }
            var duplicateSiteContentType = dependency.SiteContentTypes
                .GroupBy(value => value == null ? string.Empty : value.ContentTypeId, StringComparer.OrdinalIgnoreCase)
                .FirstOrDefault(group => string.IsNullOrWhiteSpace(group.Key) || group.Count() > 1);
            if (duplicateSiteContentType != null)
            {
                throw new InvalidDataException("List '" + dependency.Title + "' contains a missing or duplicate site content type closure node.");
            }
            foreach (var contentType in dependency.SiteContentTypes)
            {
                ContentTypeSchemaContractValidator.ValidateSnapshot(contentType);
                if (string.IsNullOrWhiteSpace(contentType.SourceScope))
                {
                    throw new InvalidDataException("Site content type '" + contentType.ContentTypeId + "' has no captured source scope.");
                }
            }
            var duplicateListContentType = dependency.ContentTypes
                .GroupBy(value => value == null ? string.Empty : value.Id, StringComparer.OrdinalIgnoreCase)
                .FirstOrDefault(group => string.IsNullOrWhiteSpace(group.Key) || group.Count() > 1);
            if (duplicateListContentType != null)
            {
                throw new InvalidDataException("List '" + dependency.Title + "' contains a missing or duplicate List content type ID '" + duplicateListContentType.Key + "'.");
            }
            var incompleteListContentType = dependency.ContentTypes.FirstOrDefault(value =>
                string.IsNullOrWhiteSpace(value.Name)
                || string.IsNullOrWhiteSpace(value.ParentId)
                || value.FieldLinks == null);
            if (incompleteListContentType != null)
            {
                throw new InvalidDataException("List content type '" + incompleteListContentType.Id + "' in List '" + dependency.Title + "' is missing its name, parent ID, or field-link collection.");
            }
            foreach (var contentType in dependency.ContentTypes)
            {
                if (contentType.FieldLinks.Any(link => link == null || link.FieldId == Guid.Empty))
                {
                    throw new InvalidDataException("List content type '" + contentType.Id + "' in List '" + dependency.Title + "' contains a null or missing-ID field link.");
                }
                var duplicateLink = contentType.FieldLinks
                    .GroupBy(link => link.FieldId)
                    .FirstOrDefault(group => group.Count() > 1);
                if (duplicateLink != null)
                {
                    throw new InvalidDataException("List content type '" + contentType.Id + "' in List '" + dependency.Title + "' contains duplicate field link '" + duplicateLink.Key.ToString("D") + "'.");
                }
            }
            var capturedSiteContentTypeIds = new HashSet<string>(dependency.SiteContentTypes.Select(value => value.ContentTypeId), StringComparer.OrdinalIgnoreCase);
            var listContentTypeIds = new HashSet<string>(dependency.ContentTypes.Select(value => value.Id), StringComparer.OrdinalIgnoreCase);
            if (dependency.UniqueContentTypeOrder.Any(value => string.IsNullOrWhiteSpace(value) || !listContentTypeIds.Contains(value))
                || dependency.UniqueContentTypeOrder.Distinct(StringComparer.OrdinalIgnoreCase).Count() != dependency.UniqueContentTypeOrder.Count)
            {
                throw new InvalidDataException("List '" + dependency.Title + "' contains an invalid or duplicate unique content type order entry.");
            }
            foreach (var contentType in dependency.ContentTypes)
            {
                if (!string.IsNullOrWhiteSpace(contentType.ParentId)
                    && !ContentTypeRuntimeCatalog.IsTargetRuntime(contentType.ParentId)
                    && !capturedSiteContentTypeIds.Contains(contentType.ParentId))
                {
                    throw new InvalidDataException("List content type '" + contentType.Id + "' has no captured custom site-content-type parent closure.");
                }
            }
            foreach (var contentType in dependency.SiteContentTypes)
            {
                if (!ContentTypeRuntimeCatalog.IsTargetRuntime(contentType.ParentContentTypeId)
                    && !capturedSiteContentTypeIds.Contains(contentType.ParentContentTypeId))
                {
                    throw new InvalidDataException("Site content type '" + contentType.ContentTypeId + "' has an uncaptured custom parent.");
                }
            }
            var duplicateView = dependency.Views.GroupBy(value => value == null ? Guid.Empty : value.Id).FirstOrDefault(group => group.Key == Guid.Empty || group.Count() > 1);
            if (duplicateView != null || dependency.Views.Any(value => !string.Equals(MigrationDigest.ComputeSha256(value.ListViewXml ?? string.Empty), value.ListViewXmlSha256, StringComparison.OrdinalIgnoreCase)))
            {
                throw new InvalidDataException("List '" + dependency.Title + "' contains missing, duplicate, or mutated view evidence.");
            }
            var renderingResources = dependency.ViewRenderingResources
                .GroupBy(value => value == null ? string.Empty : value.Id, StringComparer.Ordinal)
                .ToDictionary(group => group.Key, group => group.ToArray(), StringComparer.Ordinal);
            if (renderingResources.Any(value => string.IsNullOrWhiteSpace(value.Key) || value.Value.Length != 1))
            {
                throw new InvalidDataException("List '" + dependency.Title + "' contains a missing or duplicate View rendering-resource identity.");
            }
            foreach (var resource in dependency.ViewRenderingResources)
            {
                if (resource.Diagnostics == null)
                {
                    throw new InvalidDataException("List '" + dependency.Title + "' contains incomplete View rendering-resource evidence.");
                }
                if (resource.Availability == EvidenceAvailability.Captured
                    || resource.Availability == EvidenceAvailability.Partial
                    || resource.Artifact != null)
                {
                    if (!Uri.TryCreate(resource.SourceAbsoluteUrl, UriKind.Absolute, out var sourceUri)
                        || sourceUri.Scheme != Uri.UriSchemeHttps && sourceUri.Scheme != Uri.UriSchemeHttp)
                    {
                        throw new InvalidDataException("Captured View rendering-resource evidence has no valid HTTP(S) source URL.");
                    }
                    if (string.IsNullOrWhiteSpace(resource.SourceServerRelativeUrl))
                    {
                        throw new InvalidDataException("Captured View rendering-resource evidence has no source server-relative path.");
                    }
                    MigrationArtifactContractValidator.Validate(
                        resource.Artifact,
                        resource.ContentBase64,
                        artifactStore,
                        "View rendering resource '" + resource.SourceAbsoluteUrl + "'");
                }
            }
            foreach (var view in dependency.Views)
            {
                if (view.RenderingResourceBindings == null)
                {
                    throw new InvalidDataException("View '" + view.Id.ToString("D") + "' has a null rendering-resource binding inventory.");
                }
                var duplicateBinding = view.RenderingResourceBindings
                    .Where(value => value != null)
                    .GroupBy(value => value.SourceProperty + "\u001f" + value.OriginalReference, StringComparer.OrdinalIgnoreCase)
                    .FirstOrDefault(group => group.Count() > 1);
                if (duplicateBinding != null
                    || view.RenderingResourceBindings.Any(value => value == null
                        || string.IsNullOrWhiteSpace(value.SourceProperty)
                        || string.IsNullOrWhiteSpace(value.OriginalReference)
                        || string.IsNullOrWhiteSpace(value.ResourceId)
                        || !renderingResources.ContainsKey(value.ResourceId)))
                {
                    throw new InvalidDataException("View '" + view.Id.ToString("D") + "' contains an invalid rendering-resource binding.");
                }
            }
            var duplicateItem = dependency.Items.GroupBy(value => value == null ? 0 : value.SourceItemId).FirstOrDefault(group => group.Key <= 0 || group.Count() > 1);
            if (duplicateItem != null)
            {
                throw new InvalidDataException("List '" + dependency.Title + "' contains a missing or duplicate source item ID.");
            }
            foreach (var item in dependency.Items)
            {
                if (item.Values == null || item.Attachments == null || item.Diagnostics == null
                    || item.Values.Any(value => value == null || string.IsNullOrWhiteSpace(value.InternalName)))
                {
                    throw new InvalidDataException("List item " + item.SourceItemId + " contains a null or unnamed field value inventory entry.");
                }
                foreach (var attachment in item.Attachments)
                {
                    ValidateBinary(attachment == null ? null : attachment.Content, artifactStore, "attachment " + (attachment == null ? "<null>" : attachment.FileName));
                }
                if (item.Document != null && item.Document.Kind == ListDocumentObjectKind.File)
                {
                    ValidateInformationProtection(item.Document.InformationProtection, dependency.Title, item.SourceItemId);
                    ValidateBinary(item.Document.Content, artifactStore, "document " + item.Document.Name);
                    if (item.Document.Content != null && item.Document.Content.Artifact != null && item.Document.Length != item.Document.Content.Artifact.Length)
                    {
                        var rightsManaged = item.Document.Content.RepresentationKind
                            == ListBinaryRepresentationKind.InformationRightsManagedEnvelope;
                        var hasExplicitMismatchEvidence = rightsManaged
                            ? item.Document.Content.Diagnostics.Any(value => value != null
                                && value.StartsWith("RightsManagedEnvelopeLengthMismatch:", StringComparison.Ordinal))
                            : item.Document.Content.Availability == EvidenceAvailability.Partial
                                && item.Document.Content.Diagnostics.Any(value => value != null
                                    && value.StartsWith("DocumentMetadataLengthMismatch:", StringComparison.Ordinal));
                        if (!hasExplicitMismatchEvidence)
                        {
                            throw new InvalidDataException("Document payload length differs from captured file metadata: " + item.Document.ServerRelativeUrl);
                        }
                    }
                }
            }
        }

        private static void ValidateInformationProtection(
            ListDocumentInformationProtectionSnapshot informationProtection,
            string listTitle,
            int sourceItemId)
        {
            if (informationProtection == null)
            {
                return;
            }
            if (string.IsNullOrWhiteSpace(informationProtection.LabelId)
                || informationProtection.Diagnostics == null
                || !Enum.IsDefined(typeof(EvidenceAvailability), informationProtection.Availability))
            {
                throw new InvalidDataException(
                    "List '" + listTitle + "' item " + sourceItemId
                    + " contains incomplete document information-protection evidence.");
            }
        }

        private static void ValidateInformationRightsManagement(ListDependencySnapshot dependency)
        {
            var informationRightsManagement = dependency.InformationRightsManagement;
            if (informationRightsManagement == null)
            {
                // The contract is optional so packages captured before IRM policy
                // evidence was introduced retain their historical semantic digest.
                return;
            }
            if (informationRightsManagement.Diagnostics == null
                || !Enum.IsDefined(typeof(EvidenceAvailability), informationRightsManagement.Availability))
            {
                throw new InvalidDataException(
                    "List '" + dependency.Title + "' contains incomplete IRM policy evidence.");
            }

            if (!informationRightsManagement.IrmEnabled)
            {
                if (informationRightsManagement.Policy != null
                    || informationRightsManagement.AuthorizationEvidence != null
                    || informationRightsManagement.Availability != EvidenceAvailability.Captured)
                {
                    throw new InvalidDataException(
                        "List '" + dependency.Title + "' contains contradictory disabled IRM policy evidence.");
                }
                return;
            }

            if (informationRightsManagement.AuthorizationEvidence != null)
            {
                LiteralHttpAuthorizationEvidence.Validate(informationRightsManagement.AuthorizationEvidence);
                if (!string.Equals(
                        informationRightsManagement.AuthorizationEvidence.Operation,
                        "capture-list-irm-policy",
                        StringComparison.Ordinal)
                    || informationRightsManagement.Availability != EvidenceAvailability.Unavailable
                    || informationRightsManagement.Policy != null)
                {
                    throw new InvalidDataException(
                        "List '" + dependency.Title + "' contains contradictory IRM authorization evidence.");
                }
                return;
            }

            if (informationRightsManagement.Availability == EvidenceAvailability.Captured
                && informationRightsManagement.Policy == null)
            {
                throw new InvalidDataException(
                    "IRM-enabled List '" + dependency.Title + "' has no captured IRM policy.");
            }
            if (informationRightsManagement.Availability == EvidenceAvailability.Unavailable)
            {
                throw new InvalidDataException(
                    "IRM-enabled List '" + dependency.Title + "' is unavailable without literal HTTP 401/403 evidence.");
            }
        }

        private static void ValidateBinary(ListBinaryArtifactSnapshot binary, IMigrationArtifactStore artifactStore, string subject)
        {
            if (binary == null)
            {
                throw new InvalidDataException("Missing binary evidence record for " + subject + ".");
            }
            if (binary.Diagnostics == null)
            {
                throw new InvalidDataException("Binary evidence diagnostics are null for " + subject + ".");
            }
            if (!Enum.IsDefined(typeof(ListBinaryRepresentationKind), binary.RepresentationKind))
            {
                throw new InvalidDataException("Binary representation kind is invalid for " + subject + ".");
            }
            if (binary.LogicalContentIdentity != null
                && (string.IsNullOrWhiteSpace(binary.LogicalContentIdentity.EvidenceSource)
                    || (string.IsNullOrWhiteSpace(binary.LogicalContentIdentity.QuickXorHash)
                        && string.IsNullOrWhiteSpace(binary.LogicalContentIdentity.ContentTag))))
            {
                throw new InvalidDataException("Binary logical-content identity is incomplete for " + subject + ".");
            }
            if (binary.ArchivedContentEvidence != null)
            {
                if (binary.ArchivedContentEvidence.Count == 0
                    || binary.ArchivedContentEvidence.Any(value => value == null)
                    || binary.ArchivedContentEvidence
                        .GroupBy(
                            value => value.Operation + "\n" + value.RequestUri,
                            StringComparer.OrdinalIgnoreCase)
                        .Any(group => group.Count() > 1))
                {
                    throw new InvalidDataException("Archived-content evidence is empty, null, or duplicated for " + subject + ".");
                }
                foreach (var evidence in binary.ArchivedContentEvidence)
                {
                    LiteralHttpArchivedContentEvidence.Validate(evidence);
                }
                if (binary.Availability != EvidenceAvailability.Unavailable || binary.Artifact != null)
                {
                    throw new InvalidDataException("Archived-content evidence must describe unavailable source bytes for " + subject + ".");
                }
            }
            if (binary.Availability == EvidenceAvailability.Captured
                || (binary.Availability == EvidenceAvailability.Partial && binary.Artifact != null))
            {
                if (binary.Artifact == null)
                {
                    throw new InvalidDataException("Captured binary evidence has no artifact descriptor for " + subject + ".");
                }
                MigrationArtifact.ReadAllBytes(binary.Artifact, binary.ContentBase64, artifactStore);
            }
        }
    }
}
