using PnP.Framework.Migration.Lists.Capture;
using PnP.Framework.Migration.Lists.ContentTypes;
using PnP.Framework.Migration.Lists.Planning;
using PnP.Framework.Migration.Lists.Items;
using PnP.Framework.Migration.Schema.ContentTypes;
using System;
using System.Collections.Generic;
using System.Linq;

namespace PnP.Framework.Migration.Lists.Execution
{
    /// <summary>
    /// The admitted transaction subset of a sealed List migration plan. The
    /// package remains immutable; this scope projects only the ingredients that
    /// the page-level execution frontier marked executable.
    /// </summary>
    internal sealed class ListMaterializationExecutionScope
    {
        private readonly IDictionary<Guid, ListSelection> lists =
            new Dictionary<Guid, ListSelection>();
        private readonly ISet<string> siteContentTypes =
            new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        private readonly ISet<string> siteFields =
            new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        private readonly ISet<string> platformFeatures =
            new HashSet<string>(StringComparer.Ordinal);

        public IList<ListSelection> Lists => lists.Values
            .OrderBy(value => value.SourceListId)
            .ToList();

        public bool HasWork => lists.Count > 0
            || siteContentTypes.Count > 0
            || siteFields.Count > 0
            || platformFeatures.Count > 0;

        public void AddList(ListSelection selection)
        {
            if (selection == null || selection.SourceListId == Guid.Empty)
            {
                throw new ArgumentException("A source List execution identity is required.", nameof(selection));
            }
            if (lists.ContainsKey(selection.SourceListId))
            {
                throw new InvalidOperationException("Duplicate source List execution identity: " + selection.SourceListId.ToString("D") + ".");
            }
            lists.Add(selection.SourceListId, selection);
        }

        public void AddSiteContentType(string sourceOwnerWebUrl, string sourceContentTypeId)
        {
            siteContentTypes.Add(SiteContentTypeKey(sourceOwnerWebUrl, sourceContentTypeId));
        }

        public void AddSiteField(string sourceOwnerWebUrl, Guid sourceFieldId)
        {
            siteFields.Add(SiteFieldKey(sourceOwnerWebUrl, sourceFieldId));
        }

        public void AddPlatformFeature(Guid sourceSiteId, Guid featureId)
        {
            platformFeatures.Add(PlatformFeatureKey(sourceSiteId, featureId));
        }

        public bool IncludesList(Guid sourceListId)
        {
            return lists.ContainsKey(sourceListId);
        }

        public bool IncludesSiteContentType(ContentTypeClosureNodePlan plan)
        {
            return plan?.Schema != null
                && siteContentTypes.Contains(SiteContentTypeKey(plan.SourceOwnerWebUrl, plan.Schema.ContentTypeId));
        }

        public bool IncludesSiteField(string sourceOwnerWebUrl, Guid sourceFieldId)
        {
            return siteFields.Contains(SiteFieldKey(sourceOwnerWebUrl, sourceFieldId));
        }

        public bool IncludesPlatformFeature(Guid sourceSiteId, Guid featureId)
        {
            return platformFeatures.Contains(PlatformFeatureKey(sourceSiteId, featureId));
        }

        public ListSelection GetList(Guid sourceListId)
        {
            if (!lists.TryGetValue(sourceListId, out var selection))
            {
                throw new InvalidOperationException("Source List is outside the admitted execution scope: " + sourceListId.ToString("D") + ".");
            }
            return selection;
        }

        public ListDependencySnapshot ProjectSource(ListDependencySnapshot source)
        {
            var selection = GetList(source.SourceListId);
            var selectedFields = source.Fields
                .Where(value => value != null && selection.FieldIds.Contains(value.Id))
                .ToList();
            var selectedFieldIds = new HashSet<Guid>(selectedFields.Select(value => value.Id));
            var selectedFieldNames = new HashSet<string>(
                selectedFields.Select(value => value.InternalName),
                StringComparer.OrdinalIgnoreCase);
            var selectedContentTypes = source.ContentTypes
                .Where(value => value != null && selection.ContentTypeIds.Contains(value.Id))
                .Select(value => ProjectContentType(value, selectedFieldIds))
                .ToList();
            var selectedContentTypeIds = new HashSet<string>(
                selectedContentTypes.Select(value => value.Id),
                StringComparer.OrdinalIgnoreCase);
            var selectedItems = source.Items
                .Where(value => value != null && selection.MaterializationItemIds.Contains(value.SourceItemId))
                .Select(value => ProjectItem(
                    value,
                    selection,
                    selectedFieldNames,
                    selectedContentTypeIds))
                .ToList();
            return new ListDependencySnapshot
            {
                SchemaVersion = source.SchemaVersion,
                SourceSiteId = source.SourceSiteId,
                SourceWebId = source.SourceWebId,
                SourceWebUrl = source.SourceWebUrl,
                SourceListId = source.SourceListId,
                Title = source.Title,
                Description = source.Description,
                TemplateFeatureId = source.TemplateFeatureId,
                BaseTemplate = source.BaseTemplate,
                BaseType = source.BaseType,
                RootFolderServerRelativeUrl = source.RootFolderServerRelativeUrl,
                Hidden = source.Hidden,
                ContentTypesEnabled = source.ContentTypesEnabled,
                EnableAttachments = source.EnableAttachments,
                EnableFolderCreation = source.EnableFolderCreation,
                EnableVersioning = source.EnableVersioning,
                EnableMinorVersions = source.EnableMinorVersions,
                EnableModeration = source.EnableModeration,
                ForceCheckout = source.ForceCheckout,
                SourceItemCount = selectedItems.Count,
                Fields = selectedFields,
                ContentTypes = selectedContentTypes,
                HasExplicitUniqueContentTypeOrder = source.HasExplicitUniqueContentTypeOrder
                    && selectedContentTypeIds.Count == source.ContentTypes.Count,
                UniqueContentTypeOrder = source.UniqueContentTypeOrder
                    .Where(selectedContentTypeIds.Contains)
                    .ToList(),
                SiteContentTypes = source.SiteContentTypes
                    .Where(value => value != null
                        && siteContentTypes.Contains(SiteContentTypeKey(
                            value.SourceScope ?? value.SourceWebUrl,
                            value.ContentTypeId)))
                    .ToList(),
                Views = source.Views
                    .Where(value => value != null && selection.ViewIds.Contains(value.Id))
                    .ToList(),
                ViewRenderingResources = source.ViewRenderingResources
                    .Where(value => value != null && selection.ViewRenderingResourceIds.Contains(value.Id))
                    .ToList(),
                Items = selectedItems,
                Availability = source.Availability,
                Diagnostics = source.Diagnostics.ToList()
            };
        }

        public ListMaterializationPlan ProjectPlan(ListMaterializationPlan plan)
        {
            var selection = GetList(plan.SourceListId);
            var targetDisposition = plan.TargetProbe != null
                && plan.TargetProbe.Disposition != ListMaterializationDisposition.Block
                    ? plan.TargetProbe.Disposition
                    : plan.Disposition != ListMaterializationDisposition.Block
                        ? plan.Disposition
                        : ListMaterializationDisposition.CreateOwned;
            return new ListMaterializationPlan
            {
                SourceSiteId = plan.SourceSiteId,
                SourceWebId = plan.SourceWebId,
                SourceListId = plan.SourceListId,
                TargetWebUrl = plan.TargetWebUrl,
                TargetSiteCollectionUrl = plan.TargetSiteCollectionUrl,
                TargetWebServerRelativeUrl = plan.TargetWebServerRelativeUrl,
                PreferredTargetRootFolderServerRelativeUrl = plan.PreferredTargetRootFolderServerRelativeUrl,
                TargetRootFolderServerRelativeUrl = plan.TargetRootFolderServerRelativeUrl,
                PreferredTargetTitle = plan.PreferredTargetTitle,
                TargetTitle = plan.TargetTitle,
                OriginalIdentifier = plan.OriginalIdentifier,
                Disposition = targetDisposition,
                Fields = plan.Fields.Where(value => value != null && selection.FieldIds.Contains(value.SourceFieldId)).ToList(),
                Views = plan.Views.Where(value => value != null && selection.ViewIds.Contains(value.SourceViewId)).ToList(),
                ViewRenderingResources = plan.ViewRenderingResources
                    .Where(value => value != null && selection.ViewRenderingResourceIds.Contains(value.SourceResourceId))
                    .ToList(),
                SiteContentTypes = plan.SiteContentTypes.Where(IncludesSiteContentType).Select(ProjectContentTypeNode).ToList(),
                RequiredFeatures = plan.RequiredFeatures
                    .Where(value => value != null && IncludesPlatformFeature(plan.SourceSiteId, value.FeatureId))
                    .ToList(),
                Issues = new List<PnP.Framework.Migration.Diagnostics.MigrationIssue>(),
                TargetProbe = null,
                PlanDigest = plan.PlanDigest
            };
        }

        public ListMigrationPlanSet ProjectPlanSet(ListMigrationPlanSet planSet)
        {
            if (planSet == null)
            {
                return null;
            }
            var selectedIds = new HashSet<Guid>(lists.Keys);
            return new ListMigrationPlanSet
            {
                SchemaVersion = planSet.SchemaVersion,
                OrderedSourceListIds = planSet.OrderedSourceListIds.Where(selectedIds.Contains).ToList(),
                Lists = planSet.Lists
                    .Where(value => value != null && selectedIds.Contains(value.SourceListId))
                    .Select(ProjectPlan)
                    .ToList(),
                Issues = new List<PnP.Framework.Migration.Diagnostics.MigrationIssue>(),
                // Ownership is always sealed to the approved full plan. The execution
                // subset is represented by the page receipt/frontier, not by rewriting
                // the approved List-plan identity.
                PlanDigest = planSet.PlanDigest
            };
        }

        private static ListContentTypeSnapshot ProjectContentType(
            ListContentTypeSnapshot source,
            ISet<Guid> selectedFieldIds)
        {
            // PageIngredientPlanEvaluator has already validated that every
            // required non-executable FieldLink was explicitly released by this
            // executable Content Type Transform action. Project only retained
            // links into the runtime transaction while leaving the complete
            // source Content Type and released cache links in the sealed package.
            return new ListContentTypeSnapshot
            {
                Id = source.Id,
                Name = source.Name,
                Description = source.Description,
                Group = source.Group,
                ParentId = source.ParentId,
                Hidden = source.Hidden,
                ReadOnly = source.ReadOnly,
                Sealed = source.Sealed,
                FieldLinks = source.FieldLinks
                    .Where(value => value != null && selectedFieldIds.Contains(value.FieldId))
                    .ToList()
            };
        }

        private static ContentTypeClosureNodePlan ProjectContentTypeNode(ContentTypeClosureNodePlan source)
        {
            return new ContentTypeClosureNodePlan
            {
                SourceOwnerWebId = source.SourceOwnerWebId,
                SourceOwnerWebUrl = source.SourceOwnerWebUrl,
                TargetOwnerWebUrl = source.TargetOwnerWebUrl,
                Schema = source.Schema,
                DeferredUntilTopologyMaterialization = source.DeferredUntilTopologyMaterialization,
                TargetProbe = null,
                TargetAdmission = null,
                PlanDigest = source.PlanDigest
            };
        }

        private static ListItemSnapshot ProjectItem(
            ListItemSnapshot source,
            ListSelection selection,
            ISet<string> selectedFieldNames,
            ISet<string> selectedContentTypeIds)
        {
            var includeItem = selection.ItemIds.Contains(source.SourceItemId);
            var includeDocument = selection.DocumentItemIds.Contains(source.SourceItemId);
            return new ListItemSnapshot
            {
                SourceItemId = source.SourceItemId,
                SourceUniqueId = source.SourceUniqueId,
                Values = includeItem
                    ? source.Values.Where(value => value != null
                        && (selectedFieldNames.Contains(value.InternalName)
                            || IsSelectedContentType(value, selectedContentTypeIds)))
                        .ToList()
                    : new List<ListItemValueSnapshot>(),
                Attachments = source.Attachments
                    .Where(value => value != null && selection.IncludesAttachment(source.SourceItemId, value.FileName))
                    .ToList(),
                Document = includeDocument ? source.Document : null,
                Availability = source.Availability,
                Diagnostics = source.Diagnostics.ToList()
            };
        }

        private static bool IsSelectedContentType(
            ListItemValueSnapshot value,
            ISet<string> selectedContentTypeIds)
        {
            if (!string.Equals(value.InternalName, "ContentTypeId", StringComparison.OrdinalIgnoreCase))
            {
                return false;
            }
            var sourceContentTypeId = value.ScalarValue ?? value.RawValue;
            return !string.IsNullOrWhiteSpace(sourceContentTypeId)
                && selectedContentTypeIds.Contains(sourceContentTypeId);
        }

        private static string SiteContentTypeKey(string sourceOwnerWebUrl, string contentTypeId)
        {
            var scope = sourceOwnerWebUrl ?? string.Empty;
            if (Uri.TryCreate(scope, UriKind.Absolute, out var absolute))
            {
                scope = absolute.AbsolutePath;
            }
            scope = Uri.UnescapeDataString(scope).Replace('\\', '/').TrimEnd('/');
            return scope + "\u001f" + (contentTypeId ?? string.Empty);
        }

        private static string PlatformFeatureKey(Guid sourceSiteId, Guid featureId)
        {
            return sourceSiteId.ToString("D") + "\u001f" + featureId.ToString("D");
        }

        private static string SiteFieldKey(string sourceOwnerWebUrl, Guid fieldId)
        {
            return SiteContentTypeKey(sourceOwnerWebUrl, fieldId.ToString("D"));
        }

        internal sealed class ListSelection
        {
            public Guid SourceListId { get; set; }

            public bool IncludeListObject { get; set; }

            public ISet<Guid> FieldIds { get; } = new HashSet<Guid>();

            public ISet<string> ContentTypeIds { get; } = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            public ISet<int> ItemIds { get; } = new HashSet<int>();

            public ISet<int> DocumentItemIds { get; } = new HashSet<int>();

            public IDictionary<int, ISet<string>> AttachmentNamesByItemId { get; } =
                new Dictionary<int, ISet<string>>();

            public ISet<Guid> ViewIds { get; } = new HashSet<Guid>();

            public ISet<string> ViewRenderingResourceIds { get; } = new HashSet<string>(StringComparer.Ordinal);

            public ISet<int> ExactAttachmentInventoryItemIds { get; } = new HashSet<int>();

            public bool ExactContentTypeInventory { get; set; }

            public bool ExactItemInventory { get; set; }

            public bool HasListScopedWork => IncludeListObject
                || FieldIds.Count > 0
                || ContentTypeIds.Count > 0
                || MaterializationItemIds.Count > 0
                || ViewIds.Count > 0;

            public ISet<int> MaterializationItemIds => new HashSet<int>(
                ItemIds.Concat(DocumentItemIds).Concat(AttachmentNamesByItemId.Keys));

            public void AddAttachment(int sourceItemId, string fileName)
            {
                if (!AttachmentNamesByItemId.TryGetValue(sourceItemId, out var names))
                {
                    names = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                    AttachmentNamesByItemId.Add(sourceItemId, names);
                }
                names.Add(fileName ?? string.Empty);
            }

            public bool IncludesAttachment(int sourceItemId, string fileName)
            {
                return AttachmentNamesByItemId.TryGetValue(sourceItemId, out var names)
                    && names.Contains(fileName ?? string.Empty);
            }
        }
    }
}
