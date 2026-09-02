using PnP.Framework.Migration.Evidence;
using PnP.Framework.Migration.Lists.Capture;
using PnP.Framework.Migration.Lists.Items;
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
                || dependency.Views == null || dependency.Items == null || dependency.Diagnostics == null)
            {
                throw new InvalidDataException("A List dependency snapshot is missing identity, metadata, or an inventory collection.");
            }
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
                    ValidateBinary(item.Document.Content, artifactStore, "document " + item.Document.Name);
                    if (item.Document.Content != null && item.Document.Content.Artifact != null && item.Document.Length != item.Document.Content.Artifact.Length)
                    {
                        var hasExplicitMismatchEvidence = item.Document.Content.Availability == EvidenceAvailability.Partial
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

        private static void ValidateBinary(ListBinaryArtifactSnapshot binary, IMigrationArtifactStore artifactStore, string subject)
        {
            if (binary == null)
            {
                throw new InvalidDataException("Missing binary evidence record for " + subject + ".");
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
