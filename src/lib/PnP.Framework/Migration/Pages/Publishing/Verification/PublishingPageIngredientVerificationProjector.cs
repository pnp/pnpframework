using PnP.Framework.Migration.Lists.Planning;
using PnP.Framework.Migration.Lists.Execution;
using PnP.Framework.Migration.Pages.Fields;
using PnP.Framework.Migration.Pages.Ingredients;
using PnP.Framework.Migration.Pages.Publishing.Execution;
using PnP.Framework.Migration.Pages.Publishing.Ingredients;
using PnP.Framework.Migration.Pages.Publishing.Packaging;
using PnP.Framework.Migration.Pages.References;
using PnP.Framework.Migration.Taxonomy;
using System;
using System.Collections.Generic;
using System.Linq;

namespace PnP.Framework.Migration.Pages.Publishing.Verification
{
    internal sealed class PublishingPageIngredientVerificationEvidence
    {
        public bool StructuralMaterializersPassed { get; set; }

        public bool PageArtifactMatched { get; set; }

        public bool ContentTypeMatched { get; set; }

        public bool PublishingContentMatched { get; set; }

        public bool SecurityMatched { get; set; }

        public bool LifecycleMatched { get; set; }

        public bool TaxonomyRelationshipsMatched { get; set; }

        public bool TopologyMatched { get; set; }

        public bool DependenciesMatched { get; set; }

        public bool RuntimeVerificationRequired { get; set; }

        public IList<PageFieldImportResult> FieldResults { get; set; } = new List<PageFieldImportResult>();

        public IList<PublishingPageWebPartVerificationResult> WebPartResults { get; set; } =
            new List<PublishingPageWebPartVerificationResult>();

        public IList<ListMaterializationReceipt> ListReceipts { get; set; } =
            new List<ListMaterializationReceipt>();
    }

    internal sealed class PublishingPageIngredientVerificationSummary
    {
        public IList<string> VerifiedIngredientIds { get; set; } = new List<string>();

        public IList<string> PendingIngredientIds { get; set; } = new List<string>();

        public IList<string> FailedIngredientIds { get; set; } = new List<string>();
    }

    internal static class PublishingPageIngredientVerificationProjector
    {
        public static PublishingPageIngredientVerificationSummary Project(
            PublishingPageMigrationPackage package,
            PublishingPageExecutionScope executionScope,
            PublishingPageIngredientVerificationEvidence evidence)
        {
            if (package == null)
            {
                throw new ArgumentNullException(nameof(package));
            }
            if (executionScope == null)
            {
                throw new ArgumentNullException(nameof(executionScope));
            }
            if (evidence == null)
            {
                throw new ArgumentNullException(nameof(evidence));
            }

            var verified = new HashSet<string>(StringComparer.Ordinal);
            var pending = new HashSet<string>(StringComparer.Ordinal);
            AddCore(package, executionScope, evidence, verified, pending);
            AddStructural(executionScope, evidence, verified);
            AddTopology(executionScope, evidence, verified);
            AddLists(package, executionScope, evidence, verified);
            AddReferences(package, executionScope, evidence, verified);
            AddFields(package, executionScope, evidence, verified);
            AddTaxonomy(package, executionScope, evidence, verified);
            AddWebParts(package, executionScope, evidence, verified);

            var executable = new HashSet<string>(executionScope.ExecutableIngredientIds, StringComparer.Ordinal);
            verified.IntersectWith(executable);
            pending.IntersectWith(executable);
            var failed = new HashSet<string>(executable, StringComparer.Ordinal);
            failed.ExceptWith(verified);
            failed.ExceptWith(pending);
            return new PublishingPageIngredientVerificationSummary
            {
                VerifiedIngredientIds = verified.OrderBy(value => value, StringComparer.Ordinal).ToList(),
                PendingIngredientIds = pending.OrderBy(value => value, StringComparer.Ordinal).ToList(),
                FailedIngredientIds = failed.OrderBy(value => value, StringComparer.Ordinal).ToList()
            };
        }

        private static void AddCore(
            PublishingPageMigrationPackage package,
            PublishingPageExecutionScope scope,
            PublishingPageIngredientVerificationEvidence evidence,
            ISet<string> verified,
            ISet<string> pending)
        {
            if (scope.Runtime)
            {
                if (evidence.RuntimeVerificationRequired)
                {
                    pending.Add(PublishingPageIngredientIds.Runtime);
                }
                else
                {
                    verified.Add(PublishingPageIngredientIds.Runtime);
                }
            }
            if (scope.PageArtifact && evidence.PageArtifactMatched)
            {
                verified.Add(PublishingPageIngredientIds.PageArtifact);
            }
            if (scope.Layout && evidence.StructuralMaterializersPassed)
            {
                verified.Add(PublishingPageIngredientIds.Layout);
            }
            if (scope.ContentType && evidence.ContentTypeMatched)
            {
                verified.Add(PublishingPageIngredientIds.ContentType);
            }
            if (scope.PublishingContent && evidence.PublishingContentMatched)
            {
                verified.Add(PublishingPageIngredientIds.PublishingContent);
            }
            if (scope.Security && evidence.SecurityMatched)
            {
                verified.Add(PublishingPageIngredientIds.Security);
            }
            if (scope.Lifecycle && evidence.LifecycleMatched)
            {
                verified.Add(PublishingPageIngredientIds.Lifecycle);
            }
        }

        private static void AddStructural(
            PublishingPageExecutionScope scope,
            PublishingPageIngredientVerificationEvidence evidence,
            ISet<string> verified)
        {
            if (!evidence.StructuralMaterializersPassed)
            {
                return;
            }
            foreach (var ingredientId in scope.ExecutableIngredientIds.Where(value =>
                         value.StartsWith("layout-resource:", StringComparison.Ordinal)
                         || value.StartsWith("page-content-type-field:", StringComparison.Ordinal)
                         || value.StartsWith("site-content-type:", StringComparison.Ordinal)
                         || value.StartsWith("site-field:", StringComparison.Ordinal)
                         || value.StartsWith("platform-feature:", StringComparison.Ordinal)
                         || value.StartsWith("view-rendering-resource:", StringComparison.Ordinal)))
            {
                verified.Add(ingredientId);
            }
        }

        private static void AddTopology(
            PublishingPageExecutionScope scope,
            PublishingPageIngredientVerificationEvidence evidence,
            ISet<string> verified)
        {
            if (!evidence.TopologyMatched)
            {
                return;
            }
            foreach (var web in scope.TopologyPlan?.SiteCollections
                         .SelectMany(value => value.Webs)
                         ?? Array.Empty<PnP.Framework.Migration.Topology.WebMappingPlan>())
            {
                verified.Add(PublishingPageIngredientIds.Web(web.SourceSiteId, web.SourceWebId));
            }
        }

        private static void AddLists(
            PublishingPageMigrationPackage package,
            PublishingPageExecutionScope scope,
            PublishingPageIngredientVerificationEvidence evidence,
            ISet<string> verified)
        {
            var receiptByListId = (evidence.ListReceipts ?? Array.Empty<ListMaterializationReceipt>())
                .Where(value => value != null)
                .GroupBy(value => value.SourceListId)
                .ToDictionary(group => group.Key, group => group.First());
            var sourceByListId = (package.Snapshot.ListDependencies
                    ?? Array.Empty<PnP.Framework.Migration.Lists.Capture.ListDependencySnapshot>())
                .Where(value => value != null)
                .GroupBy(value => value.SourceListId)
                .ToDictionary(group => group.Key, group => group.First());
            foreach (var selection in scope.ListScope?.Lists ?? Array.Empty<ListMaterializationExecutionScope.ListSelection>())
            {
                if (!selection.HasListScopedWork
                    || !receiptByListId.TryGetValue(selection.SourceListId, out var receipt)
                    || !receipt.FreshReadbackPassed
                    || !sourceByListId.TryGetValue(selection.SourceListId, out var source))
                {
                    continue;
                }
                if (selection.IncludeListObject)
                {
                    verified.Add(PublishingPageIngredientIds.List(source.SourceWebId, source.SourceListId));
                }
                foreach (var fieldId in selection.FieldIds)
                {
                    verified.Add(PublishingPageIngredientIds.ListField(source.SourceWebId, source.SourceListId, fieldId));
                }
                foreach (var contentTypeId in selection.ContentTypeIds)
                {
                    verified.Add(PublishingPageIngredientIds.ListContentType(source.SourceWebId, source.SourceListId, contentTypeId));
                }
                foreach (var itemId in selection.ItemIds)
                {
                    verified.Add(PublishingPageIngredientIds.ListItem(source.SourceWebId, source.SourceListId, itemId));
                }
                foreach (var itemId in selection.DocumentItemIds)
                {
                    verified.Add(PublishingPageIngredientIds.ListDocument(source.SourceWebId, source.SourceListId, itemId));
                }
                foreach (var pair in selection.AttachmentNamesByItemId)
                {
                    foreach (var fileName in pair.Value)
                    {
                        verified.Add(PublishingPageIngredientIds.ListAttachment(
                            source.SourceWebId,
                            source.SourceListId,
                            pair.Key,
                            fileName));
                    }
                }
                foreach (var viewId in selection.ViewIds)
                {
                    verified.Add(PublishingPageIngredientIds.View(source.SourceWebId, source.SourceListId, viewId));
                }
            }
        }

        private static void AddReferences(
            PublishingPageMigrationPackage package,
            PublishingPageExecutionScope scope,
            PublishingPageIngredientVerificationEvidence evidence,
            ISet<string> verified)
        {
            foreach (var action in scope.ReferenceActions(package))
            {
                var passed = action.Disposition == PageReferenceDisposition.MaterializeAtTarget
                    ? evidence.DependenciesMatched
                    : evidence.PublishingContentMatched;
                if (passed)
                {
                    verified.Add(PublishingPageIngredientIds.Reference(action.SnapshotDependencyId));
                }
            }
        }

        private static void AddFields(
            PublishingPageMigrationPackage package,
            PublishingPageExecutionScope scope,
            PublishingPageIngredientVerificationEvidence evidence,
            ISet<string> verified)
        {
            var resultByName = (evidence.FieldResults ?? Array.Empty<PageFieldImportResult>())
                .Where(value => value != null && !string.IsNullOrWhiteSpace(value.InternalName))
                .GroupBy(value => value.InternalName, StringComparer.OrdinalIgnoreCase)
                .ToDictionary(group => group.Key, group => group.First(), StringComparer.OrdinalIgnoreCase);
            foreach (var action in scope.PageFieldActions(package))
            {
                if (resultByName.TryGetValue(action.SourceInternalName, out var result)
                    && result.Attempted
                    && result.Succeeded)
                {
                    verified.Add(PublishingPageIngredientIds.Field(action.SourceInternalName));
                }
            }
        }

        private static void AddTaxonomy(
            PublishingPageMigrationPackage package,
            PublishingPageExecutionScope scope,
            PublishingPageIngredientVerificationEvidence evidence,
            ISet<string> verified)
        {
            if (!evidence.TaxonomyRelationshipsMatched)
            {
                return;
            }
            foreach (var action in scope.TaxonomyActions(package))
            {
                verified.Add(PublishingPageIngredientIds.TaxonomyRelationship(
                    action.SourceFieldId,
                    action.SourceTermId,
                    action.SourceWssId));
            }
        }

        private static void AddWebParts(
            PublishingPageMigrationPackage package,
            PublishingPageExecutionScope scope,
            PublishingPageIngredientVerificationEvidence evidence,
            ISet<string> verified)
        {
            var passed = new HashSet<Guid>(
                (evidence.WebPartResults ?? Array.Empty<PublishingPageWebPartVerificationResult>())
                    .Where(value => value != null && value.Passed)
                    .Select(value => value.SourceWebPartId));
            foreach (var action in scope.WebPartActions(package).Where(value => passed.Contains(value.SourceWebPartId)))
            {
                verified.Add(PublishingPageIngredientIds.WebPart(action.SourceWebPartId));
            }
        }
    }
}
