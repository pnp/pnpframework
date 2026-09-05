using PnP.Framework.Migration.Lists.Execution;
using PnP.Framework.Migration.Lists.Capture;
using PnP.Framework.Migration.Lists.Planning;
using PnP.Framework.Migration.Pages.ClassicWebParts.Bindings;
using PnP.Framework.Migration.Pages.Fields;
using PnP.Framework.Migration.Pages.Ingredients;
using PnP.Framework.Migration.Pages.Publishing.Ingredients;
using PnP.Framework.Migration.Pages.Publishing.Layouts;
using PnP.Framework.Migration.Pages.Publishing.Packaging;
using PnP.Framework.Migration.Pages.References;
using PnP.Framework.Migration.Schema.Fields;
using PnP.Framework.Migration.Taxonomy;
using PnP.Framework.Migration.Topology;
using PnP.Framework.Migration.Topology.Execution;
using System;
using System.Collections.Generic;
using System.Linq;

namespace PnP.Framework.Migration.Pages.Publishing.Execution
{
    /// <summary>
    /// Immutable execution view over a sealed page package. It never changes a
    /// disposition or target identity; it exposes only transactions admitted by
    /// the dependency-aware ingredient frontier.
    /// </summary>
    internal sealed class PublishingPageExecutionScope
    {
        private readonly ISet<string> executableIngredientIds;

        private PublishingPageExecutionScope(
            PageIngredientExecutionFrontier frontier,
            ISet<string> executableIngredientIds)
        {
            Frontier = frontier;
            this.executableIngredientIds = executableIngredientIds;
        }

        public PageIngredientExecutionFrontier Frontier { get; }

        public TopologyPlan TopologyPlan { get; private set; }

        public ListMaterializationExecutionScope ListScope { get; private set; }

        public bool IsPartial => Frontier?.IsPartial == true;

        public bool HasWork => executableIngredientIds.Count > 0;

        public IList<string> ExecutableIngredientIds => executableIngredientIds
            .OrderBy(value => value, StringComparer.Ordinal)
            .ToList();

        public bool Runtime => Includes(PublishingPageIngredientIds.Runtime);

        public bool PageArtifact => Includes(PublishingPageIngredientIds.PageArtifact);

        public bool Layout => Includes(PublishingPageIngredientIds.Layout);

        public bool ContentType => Includes(PublishingPageIngredientIds.ContentType);

        public bool PublishingContent => Includes(PublishingPageIngredientIds.PublishingContent);

        public bool Security => Includes(PublishingPageIngredientIds.Security);

        public bool Lifecycle => Includes(PublishingPageIngredientIds.Lifecycle);

        public bool Includes(string ingredientId)
        {
            return !string.IsNullOrWhiteSpace(ingredientId)
                && executableIngredientIds.Contains(ingredientId);
        }

        public IList<PageFieldAction> PageFieldActions(PublishingPageMigrationPackage package)
        {
            return (package.Plan.FieldActions ?? Array.Empty<PageFieldAction>())
                .Where(value => value != null
                    && Includes(PublishingPageIngredientIds.Field(value.SourceInternalName)))
                .ToList();
        }

        public IList<TaxonomyRelationshipAction> TaxonomyActions(PublishingPageMigrationPackage package)
        {
            return (package.Plan.TaxonomyRelationshipActions ?? Array.Empty<TaxonomyRelationshipAction>())
                .Where(value => value != null
                    && Includes(PublishingPageIngredientIds.TaxonomyRelationship(
                        value.SourceFieldId,
                        value.SourceTermId,
                        value.SourceWssId)))
                .ToList();
        }

        public IList<ClassicWebPartAction> WebPartActions(PublishingPageMigrationPackage package)
        {
            return (package.Plan.WebPartActions ?? Array.Empty<ClassicWebPartAction>())
                .Where(value => value != null
                    && Includes(PublishingPageIngredientIds.WebPart(value.SourceWebPartId)))
                .ToList();
        }

        public IList<PageReferenceAction> ReferenceActions(PublishingPageMigrationPackage package)
        {
            return (package.Plan.DependencyActions ?? Array.Empty<PageReferenceAction>())
                .Where(value => value != null
                    && Includes(PublishingPageIngredientIds.Reference(value.SnapshotDependencyId)))
                .ToList();
        }

        public IList<FieldSchemaMaterializationPlan> PageContentTypeFields(
            PublishingPageMigrationPackage package)
        {
            return (package.Plan.LayoutMaterialization?.ContentTypeSchema?.Fields
                    ?? Array.Empty<FieldSchemaMaterializationPlan>())
                .Where(value => value != null
                    && Includes(PublishingPageIngredientIds.PageContentTypeField(value.FieldId)))
                .ToList();
        }

        public IList<PublishingPageLayoutResourceMaterializationPlan> LayoutResources(
            PublishingPageMigrationPackage package)
        {
            return (package.Plan.LayoutMaterialization?.ResourceMaterializations
                    ?? Array.Empty<PublishingPageLayoutResourceMaterializationPlan>())
                .Where(value => value != null
                    && Includes(PublishingPageIngredientIds.LayoutResource(value.SourceReference)))
                .ToList();
        }

        public static PublishingPageExecutionScope Create(PublishingPageMigrationPackage package)
        {
            if (package?.Plan?.ExecutionFrontier == null)
            {
                throw new ArgumentException("A sealed ingredient execution frontier is required.", nameof(package));
            }
            var executable = new HashSet<string>(
                package.Plan.ExecutionFrontier.Decisions
                    .Where(value => value != null
                        && value.State == PageIngredientExecutionState.Executable
                        && !string.IsNullOrWhiteSpace(value.IngredientId))
                    .Select(value => value.IngredientId),
                StringComparer.Ordinal);
            var result = new PublishingPageExecutionScope(package.Plan.ExecutionFrontier, executable);
            result.TopologyPlan = TopologyExecutionPlanProjector.Project(
                package.Plan.Topology,
                executable,
                PublishingPageIngredientIds.Web);
            result.ListScope = BuildListScope(package, result);
            return result;
        }

        private static ListMaterializationExecutionScope BuildListScope(
            PublishingPageMigrationPackage package,
            PublishingPageExecutionScope pageScope)
        {
            var result = new ListMaterializationExecutionScope();
            var sources = package.Snapshot.ListDependencies ?? Array.Empty<ListDependencySnapshot>();
            var plans = (package.Plan.ListMigration?.Lists ?? Array.Empty<ListMaterializationPlan>())
                .Where(value => value != null)
                .ToDictionary(value => value.SourceListId);

            foreach (var source in sources.Where(value => value != null))
            {
                plans.TryGetValue(source.SourceListId, out var listPlan);
                var selection = new ListMaterializationExecutionScope.ListSelection
                {
                    SourceListId = source.SourceListId,
                    IncludeListObject = pageScope.Includes(PublishingPageIngredientIds.List(
                        source.SourceWebId,
                        source.SourceListId))
                };
                foreach (var field in source.Fields.Where(value => value != null
                             && pageScope.Includes(PublishingPageIngredientIds.ListField(
                                 source.SourceWebId,
                                 source.SourceListId,
                                 value.Id))))
                {
                    selection.FieldIds.Add(field.Id);
                }
                foreach (var contentType in source.ContentTypes.Where(value => value != null
                             && pageScope.Includes(PublishingPageIngredientIds.ListContentType(
                                 source.SourceWebId,
                                 source.SourceListId,
                                 value.Id))))
                {
                    selection.ContentTypeIds.Add(contentType.Id);
                }
                foreach (var item in source.Items.Where(value => value != null))
                {
                    if (pageScope.Includes(PublishingPageIngredientIds.ListItem(
                        source.SourceWebId,
                        source.SourceListId,
                        item.SourceItemId)))
                    {
                        selection.ItemIds.Add(item.SourceItemId);
                    }
                    if (item.Document != null
                        && pageScope.Includes(PublishingPageIngredientIds.ListDocument(
                            source.SourceWebId,
                            source.SourceListId,
                            item.SourceItemId)))
                    {
                        selection.DocumentItemIds.Add(item.SourceItemId);
                    }
                    foreach (var attachment in item.Attachments.Where(value => value != null
                                 && pageScope.Includes(PublishingPageIngredientIds.ListAttachment(
                                     source.SourceWebId,
                                     source.SourceListId,
                                     item.SourceItemId,
                                     value.FileName))))
                    {
                        selection.AddAttachment(item.SourceItemId, attachment.FileName);
                    }
                    if (item.Attachments.All(value => value != null
                        && selection.IncludesAttachment(item.SourceItemId, value.FileName)))
                    {
                        selection.ExactAttachmentInventoryItemIds.Add(item.SourceItemId);
                    }
                }
                foreach (var view in source.Views.Where(value => value != null
                             && pageScope.Includes(PublishingPageIngredientIds.View(
                                 source.SourceWebId,
                                 source.SourceListId,
                                 value.Id))))
                {
                    selection.ViewIds.Add(view.Id);
                }
                foreach (var resource in source.ViewRenderingResources.Where(value => value != null
                             && pageScope.Includes(PublishingPageIngredientIds.ViewRenderingResource(
                                 source.SourceSiteId,
                                 value.Id))))
                {
                    selection.ViewRenderingResourceIds.Add(resource.Id);
                }
                selection.ExactContentTypeInventory = source.ContentTypes.All(value => value != null
                    && selection.ContentTypeIds.Contains(value.Id));
                selection.ExactItemInventory = source.Items.All(value => value != null
                    && selection.ItemIds.Contains(value.SourceItemId));
                if (selection.HasListScopedWork || selection.ViewRenderingResourceIds.Count > 0)
                {
                    if (listPlan == null)
                    {
                        throw new InvalidOperationException(
                            "An executable List ingredient has no typed List materialization plan: "
                            + source.SourceListId.ToString("D") + ".");
                    }
                    result.AddList(selection);
                }

                foreach (var schema in source.SiteContentTypes.Where(value => value != null))
                {
                    var scope = PublishingPageListSchemaIngredientGraphProjector.SchemaScope(schema);
                    if (pageScope.Includes(PublishingPageIngredientIds.SiteContentType(scope, schema.ContentTypeId)))
                    {
                        result.AddSiteContentType(scope, schema.ContentTypeId);
                    }
                    foreach (var field in schema.RequiredFieldClosure.Where(value => value != null
                                 && pageScope.Includes(PublishingPageIngredientIds.SiteField(scope, value.Id))))
                    {
                        result.AddSiteField(scope, field.Id);
                    }
                }

                if (listPlan != null)
                {
                    foreach (var feature in listPlan.RequiredFeatures.Where(value => value != null
                                 && pageScope.Includes(PublishingPageIngredientIds.PlatformFeature(
                                     source.SourceSiteId,
                                     value.FeatureId))))
                    {
                        result.AddPlatformFeature(source.SourceSiteId, feature.FeatureId);
                    }
                }
            }
            return result;
        }
    }
}
