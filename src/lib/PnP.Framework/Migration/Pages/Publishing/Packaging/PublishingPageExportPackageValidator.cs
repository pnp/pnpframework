using PnP.Framework.Migration.Evidence;
using PnP.Framework.Migration.Lists.Packaging;
using PnP.Framework.Migration.Packaging;
using PnP.Framework.Migration.Pages;
using PnP.Framework.Migration.Pages.Ingredients;
using PnP.Framework.Migration.Pages.Markup;
using PnP.Framework.Migration.Pages.Profiles;
using PnP.Framework.Migration.Pages.Publishing.Capture;
using PnP.Framework.Migration.Pages.Publishing.Ingredients;
using PnP.Framework.Migration.Pages.Publishing.Layouts.Packaging;
using PnP.Framework.Migration.Pages.Publishing.Profiles;
using PnP.Framework.Migration.Pages.Runtime;
using PnP.Framework.Migration.Pages.Fields;
using PnP.Framework.Migration.Pages.Fields.Taxonomy;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace PnP.Framework.Migration.Pages.Publishing.Packaging
{
    internal static class PublishingPageExportPackageValidator
    {
        public static void Validate(
            PublishingPageExportPackage package,
            IMigrationArtifactStore artifactStore)
        {
            if (package == null)
            {
                throw new InvalidDataException("The publishing-page export is empty.");
            }
            if (!string.Equals(package.SchemaVersion, PublishingPagePackageContract.ExportSchemaVersion, StringComparison.Ordinal))
            {
                throw new InvalidDataException($"Unsupported publishing-page export schema '{package.SchemaVersion}'.");
            }

            ValidateSelection(package.Selection);
            if (!string.Equals(
                    PublishingPageDigest.ComputeSelectionDigest(package.Selection),
                    package.SelectionDigest,
                    StringComparison.OrdinalIgnoreCase))
            {
                throw new InvalidDataException("The workflow selection digest does not match the package selection.");
            }

            var snapshot = package.Snapshot;
            ValidateSnapshotShape(snapshot);
            ValidatePageArtifact(snapshot.PageArtifact, snapshot.Source, artifactStore);
            ValidateRuntime(snapshot.Runtime);
            ValidateProfileSignals(snapshot.ProfileSignals);
            ValidateIngredientGraph(snapshot.IngredientGraph);
            ValidateDerivedRuntime(snapshot);
            ValidateDerivedProfileSignals(snapshot);

            var contentDigest = PublishingPageDigest.ComputeSha256(snapshot.PublishingPageContent ?? string.Empty);
            if (!string.Equals(contentDigest, snapshot.PublishingPageContentSha256, StringComparison.OrdinalIgnoreCase))
            {
                throw new InvalidDataException("The PublishingPageContent digest does not match the source HTML.");
            }

            PublishingPageLayoutPackageValidator.ValidateSnapshot(snapshot.Layout, artifactStore);
            ValidatePageFields(snapshot);
            ValidateWebParts(snapshot);
            ListDependencyPackageValidator.Validate(
                snapshot.WebParts,
                snapshot.ListWebPartBindings,
                snapshot.ListDependencies,
                snapshot.ListLookupDependencies,
                snapshot.SourceTopology,
                artifactStore);
            ValidateDependencies(snapshot);
            ValidateDerivedIngredientGraph(snapshot);

            var snapshotDigest = PublishingPageDigest.ComputeSnapshotDigest(snapshot);
            if (!string.Equals(snapshotDigest, package.SnapshotDigest, StringComparison.OrdinalIgnoreCase))
            {
                var canUseLegacyViewRenderingDigest = snapshot.ListDependencies.All(dependency =>
                    dependency != null
                    && (dependency.ViewRenderingResources?.Count ?? 0) == 0
                    && (dependency.Views ?? Array.Empty<Lists.Views.ListViewSnapshot>()).All(view =>
                        view != null && (view.RenderingResourceBindings?.Count ?? 0) == 0));
                var legacyDigest = canUseLegacyViewRenderingDigest
                    ? PublishingPageDigest.ComputeLegacySnapshotDigestWithoutViewRenderingResources(snapshot)
                    : null;
                if (!string.Equals(legacyDigest, package.SnapshotDigest, StringComparison.OrdinalIgnoreCase))
                {
                    throw new InvalidDataException("The source snapshot digest does not match the export payload.");
                }
            }
        }

        private static void ValidateSnapshotShape(PublishingPageCaptureBundle snapshot)
        {
            if (snapshot == null)
            {
                throw new InvalidDataException("The publishing-page export must contain a source snapshot.");
            }
            if (snapshot.Source == null
                || snapshot.PageArtifact == null
                || snapshot.Runtime == null
                || snapshot.IngredientGraph == null
                || snapshot.Layout == null
                || snapshot.CapturePolicy == null
                || snapshot.Security == null
                || snapshot.Lifecycle == null
                || snapshot.SourceFence == null)
            {
                throw new InvalidDataException("The source snapshot is missing identity, page artifact, runtime, ingredient graph, publishing layout, policy, security, lifecycle, or source fence evidence.");
            }
            if (snapshot.ProfileSignals == null
                || snapshot.Fields == null
                || snapshot.WebParts == null
                || snapshot.ListWebPartBindings == null
                || snapshot.ListDependencies == null
                || snapshot.ListLookupDependencies == null
                || snapshot.Dependencies == null
                || snapshot.Blockers == null
                || snapshot.Warnings == null)
            {
                throw new InvalidDataException("The source snapshot contains a null inventory collection.");
            }
        }

        private static void ValidateSelection(PublishingPageWorkflowSelection selection)
        {
            if (selection == null
                || string.IsNullOrWhiteSpace(selection.WorkflowId)
                || selection.ValidationCohort == null
                || string.IsNullOrWhiteSpace(selection.ValidationCohort.CohortId)
                || string.IsNullOrWhiteSpace(selection.ValidationCohort.PolicyVersion)
                || selection.ValidationCohort.Disposition == 0
                || selection.ValidationCohort.Reasons == null
                || !selection.ValidationCohort.Reasons.Any(value => !string.IsNullOrWhiteSpace(value)))
            {
                throw new InvalidDataException("The package must identify its workflow and validation-cohort assessment.");
            }
        }

        private static void ValidatePageArtifact(
            PageArtifactSnapshot artifact,
            PageIdentity source,
            IMigrationArtifactStore artifactStore)
        {
            if (!string.Equals(artifact.SchemaVersion, "pnp-page-artifact/v1", StringComparison.Ordinal)
                || artifact.Diagnostics == null)
            {
                throw new InvalidDataException("The page artifact has an unsupported schema or null diagnostics collection.");
            }
            if (artifact.FileUniqueId != source.FileUniqueId
                || !string.Equals(artifact.ServerRelativeUrl, source.PageServerRelativeUrl, StringComparison.OrdinalIgnoreCase))
            {
                throw new InvalidDataException("The page artifact identity does not match the captured source page identity.");
            }
            if (artifact.Availability == EvidenceAvailability.Captured
                || artifact.Availability == EvidenceAvailability.Partial)
            {
                if (artifact.Bytes == null)
                {
                    throw new InvalidDataException("Captured page artifact evidence has no byte reference.");
                }
                MigrationArtifact.ReadAllBytes(artifact.Bytes, artifact.ContentBase64, artifactStore);
            }
        }

        private static void ValidateRuntime(PageRuntimeSnapshot runtime)
        {
            if (!string.Equals(runtime.SchemaVersion, "pnp-page-runtime/v1", StringComparison.Ordinal)
                || string.IsNullOrWhiteSpace(runtime.AdapterId)
                || runtime.Diagnostics == null)
            {
                throw new InvalidDataException("The page runtime evidence has an unsupported schema or missing required value.");
            }
            var unknown = string.Equals(runtime.AdapterId, PageRuntimeAdapterIds.Unknown, StringComparison.Ordinal);
            if ((unknown && runtime.ResolutionState != PageRuntimeResolutionState.Unknown)
                || (!unknown && runtime.ResolutionState == PageRuntimeResolutionState.Unknown))
            {
                throw new InvalidDataException("The page runtime adapter and resolution state contradict each other.");
            }
        }

        private static void ValidateProfileSignals(IEnumerable<PageProfileSignal> signals)
        {
            var values = signals.ToList();
            if (values.Any(value => value == null || string.IsNullOrWhiteSpace(value.ProfileId)))
            {
                throw new InvalidDataException("Every profile signal must identify a profile.");
            }
            var duplicate = values
                .GroupBy(value => value.ProfileId + "\u001f" + value.Kind + "\u001f" + value.Subject, StringComparer.Ordinal)
                .FirstOrDefault(group => group.Count() > 1);
            if (duplicate != null)
            {
                throw new InvalidDataException("The source snapshot contains a duplicate profile signal.");
            }
        }

        private static void ValidateIngredientGraph(CanonicalPageIngredientGraph graph)
        {
            if (!string.Equals(graph.SchemaVersion, "pnp-page-ingredient-graph/v1", StringComparison.Ordinal)
                || (!string.IsNullOrWhiteSpace(graph.ProjectionVersion)
                    && !string.Equals(
                        graph.ProjectionVersion,
                        PublishingPageIngredientGraphProjector.CurrentProjectionVersion,
                        StringComparison.Ordinal)
                    && !string.Equals(
                        graph.ProjectionVersion,
                        PublishingPageIngredientGraphProjector.ProjectionVersionV6,
                        StringComparison.Ordinal)
                    && !string.Equals(
                        graph.ProjectionVersion,
                        PublishingPageIngredientGraphProjector.ProjectionVersionV5,
                        StringComparison.Ordinal)
                    && !string.Equals(
                        graph.ProjectionVersion,
                        PublishingPageIngredientGraphProjector.ProjectionVersionV4,
                        StringComparison.Ordinal)
                    && !string.Equals(
                        graph.ProjectionVersion,
                        PublishingPageIngredientGraphProjector.ProjectionVersionV3,
                        StringComparison.Ordinal)
                    && !string.Equals(
                        graph.ProjectionVersion,
                        PublishingPageIngredientGraphProjector.ProjectionVersionV2,
                        StringComparison.Ordinal))
                || graph.Nodes == null
                || graph.Edges == null)
            {
                throw new InvalidDataException("The canonical ingredient graph has an unsupported schema/projection or null collection.");
            }
            var nodes = graph.Nodes.ToList();
            var duplicateNode = nodes
                .GroupBy(value => value?.Id, StringComparer.Ordinal)
                .FirstOrDefault(group => string.IsNullOrWhiteSpace(group.Key) || group.Count() > 1);
            if (duplicateNode != null || nodes.Any(value => value == null || value.EvidenceReferences == null))
            {
                throw new InvalidDataException($"The ingredient graph contains a null, missing, or duplicate node ID '{duplicateNode?.Key}'.");
            }
            var nodeIds = new HashSet<string>(nodes.Select(value => value.Id), StringComparer.Ordinal);
            if (graph.Edges.Any(edge => edge == null))
            {
                throw new InvalidDataException("The ingredient graph contains a null dependency edge.");
            }
            var disconnectedEdge = graph.Edges.FirstOrDefault(edge => !nodeIds.Contains(edge.FromIngredientId ?? string.Empty)
                || !nodeIds.Contains(edge.ToIngredientId ?? string.Empty));
            if (disconnectedEdge != null)
            {
                var from = disconnectedEdge.FromIngredientId ?? "<null>";
                var to = disconnectedEdge.ToIngredientId ?? "<null>";
                throw new InvalidDataException(
                    $"Every ingredient edge must connect two captured graph nodes. Disconnected edge '{from}' -> '{to}'; "
                    + $"fromExists={nodeIds.Contains(from)}; toExists={nodeIds.Contains(to)}; "
                    + $"relationship={disconnectedEdge.Relationship}; requirement={disconnectedEdge.Requirement}.");
            }
            var duplicateEdge = graph.Edges
                .GroupBy(edge => edge.FromIngredientId + "\u001f" + edge.ToIngredientId + "\u001f"
                    + edge.Relationship + "\u001f" + edge.Requirement + "\u001f" + edge.Condition, StringComparer.Ordinal)
                .FirstOrDefault(group => group.Count() > 1);
            if (duplicateEdge != null)
            {
                throw new InvalidDataException("The ingredient graph contains a duplicate dependency edge.");
            }
        }

        private static void ValidatePageFields(PublishingPageCaptureBundle snapshot)
        {
            var duplicateField = snapshot.Fields
                .GroupBy(item => item?.InternalName, StringComparer.OrdinalIgnoreCase)
                .FirstOrDefault(group => string.IsNullOrWhiteSpace(group.Key) || group.Count() > 1);
            if (duplicateField != null)
            {
                throw new InvalidDataException($"The source field inventory contains a missing or duplicate internal name '{duplicateField.Key}'.");
            }

            foreach (var field in snapshot.Fields.Where(item => item != null))
            {
                if (field.StringValues == null
                    || field.LookupValues == null
                    || field.TaxonomyValues == null
                    || field.Diagnostics == null)
                {
                    throw new InvalidDataException($"Field '{field.InternalName}' contains a null value or diagnostics collection.");
                }
                if (field.Kind != PageFieldValueKind.Taxonomy
                    && field.Kind != PageFieldValueKind.TaxonomyCollection)
                {
                    continue;
                }

                var errors = PageTaxonomyRelationshipEvidence.ValidateSealedField(field);
                if (errors.Count > 0)
                {
                    throw new InvalidDataException(string.Join(" ", errors));
                }
            }
        }

        private static void ValidateWebParts(PublishingPageCaptureBundle snapshot)
        {
            foreach (var webPart in snapshot.WebParts)
            {
                if (webPart == null)
                {
                    throw new InvalidDataException("The Web Part inventory contains a null entry.");
                }
                if (!string.Equals(
                        PublishingPageDigest.ComputeSha256(webPart.ExportXml ?? string.Empty),
                        webPart.ExportSha256,
                        StringComparison.OrdinalIgnoreCase))
                {
                    throw new InvalidDataException($"Web Part export digest mismatch: {webPart.Id}");
                }
            }
        }

        private static void ValidateDependencies(PublishingPageCaptureBundle snapshot)
        {
            var duplicateDependency = snapshot.Dependencies
                .GroupBy(item => item?.Id, StringComparer.Ordinal)
                .FirstOrDefault(group => string.IsNullOrWhiteSpace(group.Key) || group.Count() > 1);
            if (duplicateDependency != null)
            {
                throw new InvalidDataException($"The dependency inventory contains a missing or duplicate ID '{duplicateDependency.Key}'.");
            }
            foreach (var dependency in snapshot.Dependencies.Where(item => !string.IsNullOrWhiteSpace(item.ContentBase64)))
            {
                byte[] payload;
                try
                {
                    payload = Convert.FromBase64String(dependency.ContentBase64);
                }
                catch (FormatException exception)
                {
                    throw new InvalidDataException($"Dependency payload is not valid Base64: {dependency.Id}", exception);
                }
                if (payload.LongLength != dependency.ContentLength
                    || !string.Equals(
                        PublishingPageDigest.ComputeSha256(payload),
                        dependency.ContentSha256,
                        StringComparison.OrdinalIgnoreCase))
                {
                    throw new InvalidDataException($"Dependency payload length or digest mismatch: {dependency.Id}");
                }
            }
        }

        private static void ValidateDerivedRuntime(PublishingPageCaptureBundle snapshot)
        {
            var expected = PageRuntimeResolver.Resolve(
                snapshot.PageArtifact,
                snapshot.Layout.PageDirective,
                snapshot.Source.ContentTypeId);
            if (!PublishingPageValidationCanonical.Equals(expected, snapshot.Runtime))
            {
                throw new InvalidDataException("The sealed page runtime does not match CLR/content-type evidence in the source snapshot.");
            }
        }

        private static void ValidateDerivedProfileSignals(PublishingPageCaptureBundle snapshot)
        {
            var expected = PublishingPageProfileSignalProjector.Project(
                snapshot.Source,
                snapshot.Layout,
                snapshot.Fields);
            if (!PublishingPageValidationCanonical.Equals(expected, snapshot.ProfileSignals))
            {
                throw new InvalidDataException("The sealed page profile signals do not match the typed source evidence.");
            }
        }

        private static void ValidateDerivedIngredientGraph(PublishingPageCaptureBundle snapshot)
        {
            if (!string.IsNullOrWhiteSpace(snapshot.IngredientGraph.ProjectionVersion))
            {
                var expected = PublishingPageIngredientGraphProjector.ProjectForVersion(
                    snapshot,
                    snapshot.IngredientGraph.ProjectionVersion);
                if (!PublishingPageValidationCanonical.Equals(expected, snapshot.IngredientGraph))
                {
                    throw new InvalidDataException("The sealed canonical ingredient graph does not match the typed source evidence for its declared projection version.");
                }
                return;
            }

            // Exports captured before projection versioning remain immutable. Validate them
            // against the legacy semantics that originally produced the sealed graph. During
            // the short development transition, also accept current semantics without a stamp;
            // new captures always write CurrentProjectionVersion.
            var legacy = PublishingPageIngredientGraphProjector.ProjectLegacy(snapshot);
            if (PublishingPageValidationCanonical.Equals(legacy, snapshot.IngredientGraph))
            {
                return;
            }
            var transitional = PublishingPageIngredientGraphProjector.ProjectCurrentUnversioned(snapshot);
            if (!PublishingPageValidationCanonical.Equals(transitional, snapshot.IngredientGraph))
            {
                throw new InvalidDataException("The sealed canonical ingredient graph does not match any supported typed-evidence projection.");
            }
        }
    }
}
