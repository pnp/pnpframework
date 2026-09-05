using PnP.Framework.Migration.Diagnostics;
using PnP.Framework.Migration.Evidence;
using PnP.Framework.Migration.Lists.Capture;
using PnP.Framework.Migration.Lists.Fields;
using PnP.Framework.Migration.Lists.Items;
using PnP.Framework.Migration.Lists.Views;
using PnP.Framework.Migration.Packaging;
using PnP.Framework.Migration.Schema.Fields;
using PnP.Framework.Migration.Schema.ContentTypes;
using PnP.Framework.Migration.Features;
using PnP.Framework.Migration.Taxonomy;
using PnP.Framework.Migration.Topology;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;

namespace PnP.Framework.Migration.Lists.Planning
{
    public static class ListMigrationPlanFactory
    {
        private static readonly HashSet<int> SupportedTemplates = new HashSet<int> { 100, 101, 103, 170, 851 };

        private static readonly HashSet<string> SupportedScalarTypes = new HashSet<string>(
            new[] { "Text", "Note", "Boolean", "Number", "Currency", "Integer", "DateTime", "Guid", "URL", "Choice", "MultiChoice" },
            StringComparer.OrdinalIgnoreCase);

        public static ListMigrationPlanSet Create(
            IEnumerable<ListDependencySnapshot> dependencies,
            IEnumerable<ListLookupDependency> lookupDependencies,
            TopologyPlan topology,
            IEnumerable<TaxonomyTargetMapping> taxonomyMappings,
            IEnumerable<ListTargetOverride> targetOverrides)
        {
            if (dependencies == null)
            {
                throw new ArgumentNullException(nameof(dependencies));
            }
            if (topology == null)
            {
                throw new ArgumentNullException(nameof(topology));
            }

            var sources = dependencies.ToArray();
            var edges = (lookupDependencies ?? Enumerable.Empty<ListLookupDependency>()).ToArray();
            var order = ListLookupDependencyGraph.Order(sources.Select(value => value.SourceListId), edges);
            var issues = order.Issues.ToList();
            var webMappings = topology.SiteCollections.SelectMany(value => value.Webs).ToDictionary(value => value.SourceWebId);
            var overrides = (targetOverrides ?? Enumerable.Empty<ListTargetOverride>())
                .GroupBy(value => value.SourceListId).ToDictionary(group => group.Key, group => group.ToArray());
            foreach (var duplicate in overrides.Where(value => value.Value.Length != 1))
            {
                issues.Add(Issue("DuplicateListTargetOverride", "list:" + duplicate.Key.ToString("D"), "A source List has more than one target override."));
            }

            var plans = new List<ListMaterializationPlan>();
            foreach (var source in sources.OrderBy(value => IndexOf(order.OrderedSourceListIds, value.SourceListId)))
            {
                WebMappingPlan owner;
                if (!webMappings.TryGetValue(source.SourceWebId, out owner))
                {
                    issues.Add(Issue("SourceListOwnerMappingUnavailable", "list:" + source.SourceListId.ToString("D"), "The source List owner Web is absent from the topology plan."));
                    continue;
                }
                ListTargetOverride[] candidates;
                var targetOverride = overrides.TryGetValue(source.SourceListId, out candidates) && candidates.Length == 1 ? candidates[0] : null;
                plans.Add(CreateListPlan(source, owner, topology, taxonomyMappings, targetOverride));
            }

            var result = new ListMigrationPlanSet
            {
                OrderedSourceListIds = order.OrderedSourceListIds.ToList(),
                Lists = plans,
                Issues = issues.OrderBy(value => value.Code, StringComparer.Ordinal).ThenBy(value => value.Subject, StringComparer.Ordinal).ToList()
            };
            result.PlanDigest = ComputeSetDigest(result);
            return result;
        }

        public static string ComputePlanDigest(ListMaterializationPlan plan)
        {
            var value = plan.PlanDigest;
            var probe = plan.TargetProbe;
            var disposition = plan.Disposition;
            var contentTypeStates = plan.SiteContentTypes.Select(node => new
            {
                Node = node,
                node.TargetProbe,
                node.TargetAdmission,
                node.DeferredUntilTopologyMaterialization
            }).ToArray();
            var featureStates = plan.RequiredFeatures.Select(feature => new
            {
                Feature = feature,
                feature.TargetProbe
            }).ToArray();
            plan.PlanDigest = null;
            plan.TargetProbe = null;
            plan.Disposition = ListMaterializationDisposition.CreateOwned;
            foreach (var state in contentTypeStates)
            {
                state.Node.TargetProbe = null;
                state.Node.TargetAdmission = null;
                state.Node.DeferredUntilTopologyMaterialization = false;
            }
            foreach (var state in featureStates)
            {
                state.Feature.TargetProbe = null;
            }
            try
            {
                return MigrationDigest.ComputeSha256(MigrationContractSerializer.SerializeCanonical(plan));
            }
            finally
            {
                plan.PlanDigest = value;
                plan.TargetProbe = probe;
                plan.Disposition = disposition;
                foreach (var state in contentTypeStates)
                {
                    state.Node.TargetProbe = state.TargetProbe;
                    state.Node.TargetAdmission = state.TargetAdmission;
                    state.Node.DeferredUntilTopologyMaterialization = state.DeferredUntilTopologyMaterialization;
                }
                foreach (var state in featureStates)
                {
                    state.Feature.TargetProbe = state.TargetProbe;
                }
            }
        }

        public static void SealTargetAnalysis(ListMigrationPlanSet plan)
        {
            if (plan == null)
            {
                throw new ArgumentNullException(nameof(plan));
            }
            foreach (var list in plan.Lists)
            {
                list.Disposition = list.TargetProbe == null || !list.TargetProbe.IsAdmitted
                    || list.SiteContentTypes.Any(value => !value.IsExecutable)
                    || list.RequiredFeatures.Any(value => !value.IsExecutable)
                    ? ListMaterializationDisposition.Block
                    : list.TargetProbe.Disposition;
            }
            plan.PlanDigest = ComputeSetDigest(plan);
        }

        public static string ComputeSetDigest(ListMigrationPlanSet plan)
        {
            var value = plan.PlanDigest;
            plan.PlanDigest = null;
            try
            {
                return MigrationDigest.ComputeSha256(MigrationContractSerializer.SerializeCanonical(plan));
            }
            finally
            {
                plan.PlanDigest = value;
            }
        }

        private static ListMaterializationPlan CreateListPlan(
            ListDependencySnapshot source,
            WebMappingPlan owner,
            TopologyPlan topology,
            IEnumerable<TaxonomyTargetMapping> taxonomyMappings,
            ListTargetOverride targetOverride)
        {
            var issues = new List<MigrationIssue>();
            if (source.Availability == EvidenceAvailability.Unavailable || source.Availability == EvidenceAvailability.Conflict)
            {
                issues.Add(Issue("ListEvidenceUnavailable", "list:" + source.SourceListId.ToString("D"), "Source List evidence is unavailable or conflicting."));
            }
            if (!SupportedTemplates.Contains(source.BaseTemplate))
            {
                issues.Add(Issue("UnsupportedListTemplate", "list:" + source.SourceListId.ToString("D"), "List template " + source.BaseTemplate + " (" + source.BaseType + ") is not implemented."));
            }
            if (source.Items.Count != source.SourceItemCount)
            {
                issues.Add(Issue("ListItemCaptureIncomplete", "list:" + source.SourceListId.ToString("D"), "Source ItemCount is " + source.SourceItemCount + ", but " + source.Items.Count + " item snapshots were captured."));
            }
            foreach (var attachment in source.Items.SelectMany(value => value.Attachments))
            {
                if (!HasReplayableBinary(attachment.Content))
                {
                    issues.Add(IsArchivedContent(attachment.Content)
                        ? Issue("ListBinaryContentArchived", "attachment:" + attachment.ServerRelativeUrl,
                            "The source attachment is stored in Microsoft 365 Archive. Reactivate it and perform a fresh capture before exact materialization.")
                        : Issue("ListBinaryEvidenceUnavailable", "attachment:" + attachment.ServerRelativeUrl,
                            "Exact attachment bytes are required before materialization."));
                }
            }
            foreach (var document in source.Items.Where(value => value.Document != null && value.Document.Kind == ListDocumentObjectKind.File).Select(value => value.Document))
            {
                if (!HasReplayableBinary(document.Content))
                {
                    issues.Add(IsArchivedContent(document.Content)
                        ? Issue("ListBinaryContentArchived", "document:" + document.ServerRelativeUrl,
                            "The source document is stored in Microsoft 365 Archive. Reactivate it and perform a fresh capture before exact materialization.")
                        : Issue("ListBinaryEvidenceUnavailable", "document:" + document.ServerRelativeUrl,
                            "Exact document bytes are required before materialization."));
                }
            }

            var contentTypeClosure = ContentTypeClosurePlanner.Create(source.SiteContentTypes, topology, taxonomyMappings);
            foreach (var issue in contentTypeClosure.Issues)
            {
                issues.Add(issue);
            }
            var capturedSiteContentTypeIds = new HashSet<string>(contentTypeClosure.Nodes.Select(value => value.Schema.ContentTypeId), StringComparer.OrdinalIgnoreCase);
            foreach (var contentType in source.ContentTypes.Where(value => !string.IsNullOrWhiteSpace(value.ParentId)
                && !ContentTypeRuntimeCatalog.IsTargetRuntime(value.ParentId)
                && !capturedSiteContentTypeIds.Contains(value.ParentId)))
            {
                issues.Add(Issue("CustomListContentTypeClosureUnavailable", "content-type:" + contentType.Id,
                    "The List uses custom content type '" + contentType.Name + "', but its exact site-content-type parent closure is missing."));
            }
            var requiredFeatures = ContentTypeRuntimeCatalog.CreateFeatureRequirements(
                source.ContentTypes.Select(value => value.ParentId),
                source.SiteContentTypes,
                topology.SiteCollections.Single(value => value.SourceSiteId == source.SourceSiteId).TargetSiteCollectionUrl);
            var siteMapping = topology.SiteCollections.Single(value => value.SourceSiteId == source.SourceSiteId);

            var fieldOrder = ListCalculatedFieldOrder.Order(source.Fields.Select(field => CreateFieldPlan(source, field, taxonomyMappings, issues)));
            var fieldPlans = fieldOrder.Fields;
            if (fieldOrder.CycleFields.Count > 0)
            {
                issues.Add(Issue("CalculatedFieldDependencyCycle", "list:" + source.SourceListId.ToString("D"),
                    "Calculated fields contain a dependency cycle: " + string.Join(", ", fieldOrder.CycleFields) + "."));
            }
            var renderingResourcePlans = source.ViewRenderingResources
                .Select(resource => CreateViewRenderingResourcePlan(resource, siteMapping, issues))
                .ToList();
            var renderingResourcePlansById = renderingResourcePlans
                .GroupBy(value => value.SourceResourceId, StringComparer.Ordinal)
                .ToDictionary(group => group.Key, group => group.First(), StringComparer.Ordinal);
            var viewPlans = source.Views.Select(view =>
            {
                var customRendering = IsCustomRenderingReference(view.JsLink) || IsCustomRenderingReference(view.XslLink);
                var unsupportedCustomXsl = view.RenderingResourceBindings?.Any(binding => binding != null
                    && string.Equals(binding.SourceProperty, "XslLink", StringComparison.OrdinalIgnoreCase)) == true;
                var resourceClosureMissing = customRendering
                    && (unsupportedCustomXsl
                        || view.RenderingResourceBindings == null
                        || view.RenderingResourceBindings.Count == 0
                        || view.RenderingResourceBindings.Any(binding => binding == null
                            || !renderingResourcePlansById.TryGetValue(binding.ResourceId ?? string.Empty, out var resourcePlan)
                            || !resourcePlan.IsExecutable));
                return new ListViewMaterializationPlan
                {
                    SourceViewId = view.Id,
                    Title = view.Title,
                    SourceServerRelativeUrl = view.ServerRelativeUrl,
                    Disposition = view.Availability == EvidenceAvailability.Unavailable || view.Availability == EvidenceAvailability.Conflict
                            || resourceClosureMissing
                        ? ListViewMaterializationDisposition.Block
                        : view.PersonalView
                            ? ListViewMaterializationDisposition.SkipPersonal
                            : view.IsPageBound
                                ? ListViewMaterializationDisposition.CreateOrReuseWebPartView
                                : ListViewMaterializationDisposition.CreateOrReuseOwnedPublicView,
                    Source = view,
                    Reason = resourceClosureMissing
                        ? "One or more custom View rendering resources have no exact captured and materializable dependency."
                        : view.PersonalView
                            ? "Personal views are user-scoped and remain evidence-only."
                            : view.IsPageBound
                                ? "The approved Web Part rewrite replays this embedded view with a target runtime view ID."
                                : "Create or exactly reuse the shared view inside the migration-owned target List."
                };
            }).ToList();
            foreach (var blockedView in viewPlans.Where(value => value.Disposition == ListViewMaterializationDisposition.Block))
            {
                issues.Add(Issue("ViewEvidenceUnavailable", "view:" + blockedView.SourceViewId.ToString("D"), "View schema evidence is unavailable or conflicting."));
            }
            foreach (var customRenderingView in viewPlans.Where(value => value.Disposition == ListViewMaterializationDisposition.Block
                && (IsCustomRenderingReference(value.Source.JsLink) || IsCustomRenderingReference(value.Source.XslLink))))
            {
                issues.Add(Issue("ViewRenderingResourceUnavailable", "view:" + customRenderingView.SourceViewId.ToString("D"),
                    "One or more custom JSLink/XslLink dependencies lack exact captured and materializable resource evidence."));
            }

            var targetPath = string.IsNullOrWhiteSpace(targetOverride == null ? null : targetOverride.TargetRootFolderServerRelativeUrl)
                ? TopologyPlanner.MapWebOwnedServerRelativePath(source.RootFolderServerRelativeUrl, owner.SourceServerRelativeUrl, owner.TargetServerRelativeUrl)
                : NormalizeTargetPath(targetOverride.TargetRootFolderServerRelativeUrl, owner.TargetServerRelativeUrl);
            var plan = new ListMaterializationPlan
            {
                SourceSiteId = source.SourceSiteId,
                SourceWebId = source.SourceWebId,
                SourceListId = source.SourceListId,
                TargetWebUrl = owner.TargetWebUrl,
                TargetSiteCollectionUrl = topology.SiteCollections.Single(value => value.SourceSiteId == source.SourceSiteId).TargetSiteCollectionUrl,
                TargetWebServerRelativeUrl = owner.TargetServerRelativeUrl,
                PreferredTargetRootFolderServerRelativeUrl = targetPath,
                TargetRootFolderServerRelativeUrl = targetPath,
                PreferredTargetTitle = string.IsNullOrWhiteSpace(targetOverride == null ? null : targetOverride.TargetTitle) ? source.Title : targetOverride.TargetTitle,
                TargetTitle = string.IsNullOrWhiteSpace(targetOverride == null ? null : targetOverride.TargetTitle) ? source.Title : targetOverride.TargetTitle,
                OriginalIdentifier = "urn:pnp:spo-list:v1:" + source.SourceSiteId.ToString("D") + ":" + source.SourceWebId.ToString("D") + ":" + source.SourceListId.ToString("D"),
                Disposition = issues.Any(value => value.Severity == MigrationIssueSeverity.Blocker) ? ListMaterializationDisposition.Block : ListMaterializationDisposition.CreateOwned,
                Fields = fieldPlans,
                Views = viewPlans,
                ViewRenderingResources = renderingResourcePlans,
                SiteContentTypes = contentTypeClosure.Nodes,
                RequiredFeatures = requiredFeatures,
                Issues = issues.OrderBy(value => value.Code, StringComparer.Ordinal).ThenBy(value => value.Subject, StringComparer.Ordinal).ToList()
            };
            plan.PlanDigest = ComputePlanDigest(plan);
            return plan;
        }

        private static ListViewRenderingResourceMaterializationPlan CreateViewRenderingResourcePlan(
            ListViewRenderingResourceSnapshot source,
            SiteCollectionMappingPlan siteMapping,
            ICollection<MigrationIssue> issues)
        {
            var plan = new ListViewRenderingResourceMaterializationPlan
            {
                SourceResourceId = source?.Id,
                Kind = source == null ? ListViewRenderingResourceKind.Other : source.Kind,
                SourceAbsoluteUrl = source?.SourceAbsoluteUrl,
                SourceServerRelativeUrl = source?.SourceServerRelativeUrl,
                SourceArtifact = source?.Artifact,
                SourceContentBase64 = source?.ContentBase64,
                Disposition = ListViewRenderingResourceMaterializationDisposition.Block
            };
            if (source == null || string.IsNullOrWhiteSpace(source.SourceServerRelativeUrl))
            {
                plan.Reason = "The View rendering-resource identity is absent from the sealed snapshot.";
                issues.Add(Issue("ViewRenderingResourceUnavailable", "view-rendering-resource:" + (source?.Id ?? "missing"), plan.Reason));
                return plan;
            }

            var owner = (siteMapping.Webs ?? Array.Empty<WebMappingPlan>())
                .Where(value => value != null
                    && !string.IsNullOrWhiteSpace(value.SourceServerRelativeUrl)
                    && IsWithin(source.SourceServerRelativeUrl, value.SourceServerRelativeUrl))
                .OrderByDescending(value => value.SourceServerRelativeUrl.Length)
                .FirstOrDefault();
            if (owner == null)
            {
                plan.Reason = "The View rendering-resource owner Web is absent from the reviewed topology plan.";
                issues.Add(Issue("ViewRenderingResourceOwnerMappingUnavailable", "view-rendering-resource:" + source.Id, plan.Reason));
                return plan;
            }

            var relativePath = source.SourceServerRelativeUrl.Substring(owner.SourceServerRelativeUrl.TrimEnd('/').Length).TrimStart('/');
            if (!IsReviewedAssetPath(relativePath))
            {
                plan.Reason = "Only SiteAssets and Style Library View rendering resources have a reviewed exact-path materializer.";
                issues.Add(Issue("ViewRenderingResourcePathUnsupported", "view-rendering-resource:" + source.Id, plan.Reason));
                return plan;
            }

            plan.TargetServerRelativeUrl = TopologyPlanner.MapWebOwnedServerRelativePath(
                source.SourceServerRelativeUrl,
                owner.SourceServerRelativeUrl,
                owner.TargetServerRelativeUrl);
            var targetAuthority = new Uri(siteMapping.TargetSiteCollectionUrl).GetLeftPart(UriPartial.Authority);
            plan.TargetAbsoluteUrl = new Uri(new Uri(targetAuthority + "/"), plan.TargetServerRelativeUrl.TrimStart('/')).AbsoluteUri;
            if (source.Availability == EvidenceAvailability.Unavailable
                || source.Availability == EvidenceAvailability.Conflict
                || source.Artifact == null)
            {
                plan.Disposition = ListViewRenderingResourceMaterializationDisposition.PreserveReferenceOnly;
                plan.Reason = "Preserve the captured JSLink/XslLink relationship at the mapped path without creating resource bytes because no exact readable payload exists in the sealed snapshot. This is an explicit lossy substitute, not an authorization block.";
                return plan;
            }
            plan.Disposition = ListViewRenderingResourceMaterializationDisposition.CreateOrReuseExact;
            plan.Reason = "Copy or exactly reuse the sealed View rendering resource at the same mapped Web-relative path.";
            return plan;
        }

        internal static bool HasReplayableBinary(ListBinaryArtifactSnapshot binary)
        {
            return binary?.Artifact != null
                && (binary.Availability == EvidenceAvailability.Captured
                    || binary.Availability == EvidenceAvailability.Partial);
        }

        internal static bool IsArchivedContent(ListBinaryArtifactSnapshot binary)
        {
            return binary?.ArchivedContentEvidence != null
                && binary.ArchivedContentEvidence.Count > 0;
        }

        internal static bool IsRightsManagedEnvelope(ListBinaryArtifactSnapshot binary)
        {
            return binary?.RepresentationKind
                == ListBinaryRepresentationKind.InformationRightsManagedEnvelope;
        }

        internal static bool IsUnclassifiedBinary(ListBinaryArtifactSnapshot binary)
        {
            return binary != null
                && binary.RepresentationKind == ListBinaryRepresentationKind.Unclassified;
        }

        private static ListFieldMaterializationPlan CreateFieldPlan(
            ListDependencySnapshot source,
            ListFieldSnapshot field,
            IEnumerable<TaxonomyTargetMapping> taxonomyMappings,
            ICollection<MigrationIssue> issues)
        {
            if (field.Availability == EvidenceAvailability.Unavailable || field.Availability == EvidenceAvailability.Conflict)
            {
                return Block(field, issues, "ListFieldEvidenceUnavailable", "Field schema evidence is unavailable or conflicting.");
            }
            var hasValue = HasBusinessValue(source, field.InternalName);
            var requiredBySurface = source.Views.Any(view => view.ViewFields.Contains(field.InternalName, StringComparer.OrdinalIgnoreCase))
                || source.ContentTypes.Any(contentType => contentType.FieldLinks.Any(link => link.FieldId == field.Id));
            if (ReviewedListRuntimeFieldCatalog.IsSnapshotOnly(field))
            {
                return Plan(
                    field,
                    ListFieldMaterializationDisposition.EvidenceOnly,
                    null,
                    "SharePoint owns this sealed cache or derived field. Preserve its captured schema and value in the immutable snapshot, but do not create the field or replay its source value.");
            }
            var targetRuntime = field.FromBaseType
                || FieldOwnershipClassifier.IsTargetRuntime(field.Id, field.SchemaXml)
                || IsListTemplateRuntimeField(source, field);
            if (targetRuntime)
            {
                var copyValue = hasValue && !field.ReadOnly && SupportedScalarTypes.Contains(field.TypeAsString);
                if (!requiredBySurface && !copyValue)
                {
                    return Plan(
                        field,
                        ListFieldMaterializationDisposition.EvidenceOnly,
                        null,
                        "The SharePoint-owned field has no retained View or Content Type consumer and no supported writable value; preserve its complete source evidence without requiring tenant-specific runtime schema.");
                }
                return Plan(field,
                    copyValue
                        ? ListFieldMaterializationDisposition.RequireTargetRuntimeAndCopyValue
                        : ListFieldMaterializationDisposition.RequireTargetRuntime,
                    null,
                    copyValue
                        ? "The target List template owns this writable runtime field identity; copy its recognized current value."
                        : "The target List template owns this runtime field identity; SharePoint-owned/read-only values are not replayed.");
            }
            if (field.TypeAsString.StartsWith("Lookup", StringComparison.OrdinalIgnoreCase))
            {
                if (!field.SourceLookupListId.HasValue)
                {
                    return Block(field, issues, "LookupMappingUnavailable", "Lookup field schema has no source List identity.");
                }
                return Plan(field, ListFieldMaterializationDisposition.MapLookup, null,
                    "Create or reuse the lookup field after its dependency List and source-to-target item ID catalog exist.");
            }
            if (field.TypeAsString.StartsWith("TaxonomyFieldType", StringComparison.OrdinalIgnoreCase))
            {
                var mapping = field.Taxonomy == null ? null : (taxonomyMappings ?? Enumerable.Empty<TaxonomyTargetMapping>()).SingleOrDefault(value =>
                    value.SourceTermStoreId == field.Taxonomy.SourceTermStoreId && value.SourceTermSetId == field.Taxonomy.SourceTermSetId);
                if (field.Taxonomy == null || mapping == null)
                {
                    return Block(field, issues, "TaxonomyMappingUnavailable", "Taxonomy field requires a reviewed target term store and term set mapping.");
                }
                var schema = FieldSchemaCanonicalizer.RewriteForTarget(field.SchemaXml, mapping.TargetTermStoreId, mapping.TargetTermSetId, field.Taxonomy.HiddenTextFieldId);
                return Plan(field, ListFieldMaterializationDisposition.MapTaxonomy, schema,
                    "Create or reuse the taxonomy field with the selected target store/set after taxonomy-asset admission; target SharePoint allocates WssIds and fresh readback verifies the binding.");
            }
            if (string.Equals(field.TypeAsString, "User", StringComparison.OrdinalIgnoreCase)
                || string.Equals(field.TypeAsString, "UserMulti", StringComparison.OrdinalIgnoreCase))
            {
                if (hasValue)
                {
                    return Block(field, issues, "PrincipalMappingUnavailable", "Non-empty User field values require an explicit principal mapping and are never guessed across sites or tenants.");
                }
                return Plan(field, ListFieldMaterializationDisposition.CreateOrReuseOwnedSchemaOnly,
                    FieldSchemaCanonicalizer.RewriteForTarget(field.SchemaXml), "Create the empty source-owned User field schema without inventing principal values.");
            }
            if ((field.ReadOnly || field.Sealed) && (field.TypeAsString.StartsWith("Calculated", StringComparison.OrdinalIgnoreCase)
                || field.TypeAsString.StartsWith("Computed", StringComparison.OrdinalIgnoreCase)))
            {
                return Plan(field, ListFieldMaterializationDisposition.CreateOrReuseOwnedCalculated,
                    FieldSchemaCanonicalizer.RewriteForTarget(field.SchemaXml), "Create or reuse the source-owned calculated field; target SharePoint computes its values.");
            }
            if (!field.ReadOnly && !field.Sealed && SupportedScalarTypes.Contains(field.TypeAsString))
            {
                return Plan(field, ListFieldMaterializationDisposition.CreateOrReuseOwnedAndCopyValue,
                    FieldSchemaCanonicalizer.RewriteForTarget(field.SchemaXml), "Create or reuse the source-owned field and copy recognized values.");
            }
            if (!hasValue && !requiredBySurface)
            {
                return Plan(field, ListFieldMaterializationDisposition.EvidenceOnly, null,
                    "The unsupported field is retained in the snapshot but has no captured value or rendered/schema consumer.");
            }
            return Block(field, issues, "ListFieldMaterializerUnavailable", "Field type '" + field.TypeAsString + "' is used by captured content but has no reviewed target materializer.");
        }

        private static ListFieldMaterializationPlan Plan(ListFieldSnapshot field, ListFieldMaterializationDisposition disposition, string targetSchema, string reason)
        {
            return new ListFieldMaterializationPlan
            {
                SourceFieldId = field.Id,
                InternalName = field.InternalName,
                Title = field.Title,
                TypeAsString = field.TypeAsString,
                Disposition = disposition,
                SourceSchemaXml = field.SchemaXml,
                TargetSchemaXml = targetSchema,
                SourcePortableSchemaSha256 = field.PortableSchemaSha256,
                TargetPortableSchemaSha256 = targetSchema == null ? null : FieldSchemaCanonicalizer.PortableDigest(targetSchema),
                SourceLookupWebId = field.SourceLookupWebId,
                SourceLookupListId = field.SourceLookupListId,
                LookupField = field.LookupField,
                Reason = reason
            };
        }

        private static ListFieldMaterializationPlan Block(ListFieldSnapshot field, ICollection<MigrationIssue> issues, string code, string reason)
        {
            issues.Add(Issue(code, "field:" + field.InternalName + ":" + field.Id.ToString("D"), reason));
            return Plan(field, ListFieldMaterializationDisposition.Block, null, reason);
        }

        private static bool HasBusinessValue(ListDependencySnapshot source, string internalName)
        {
            return source.Items.Any(item => item.Values.Any(value => string.Equals(value.InternalName, internalName, StringComparison.OrdinalIgnoreCase)
                && value.Kind != ListItemValueKind.Null));
        }

        private static bool HasPlatformSource(string schemaXml)
        {
            try
            {
                var source = XDocument.Parse(schemaXml).Root?.Attribute("SourceID")?.Value;
                return source != null && source.StartsWith("http://schemas.microsoft.com/sharepoint/", StringComparison.OrdinalIgnoreCase);
            }
            catch (System.Xml.XmlException)
            {
                return false;
            }
        }

        private static bool IsListTemplateRuntimeField(ListDependencySnapshot source, ListFieldSnapshot field)
        {
            if (source == null || field == null || (!field.ReadOnly && !field.Sealed) || string.IsNullOrWhiteSpace(field.SchemaXml))
            {
                return false;
            }

            try
            {
                var sourceId = XDocument.Parse(field.SchemaXml).Root?.Attribute("SourceID")?.Value;
                return Guid.TryParse(sourceId?.Trim('{', '}'), out var ownerId) && ownerId == source.SourceListId;
            }
            catch (System.Xml.XmlException)
            {
                return false;
            }
        }

        private static bool IsCustomRenderingReference(string value)
        {
            if (string.IsNullOrWhiteSpace(value))
            {
                return false;
            }
            return value.IndexOf('/') >= 0 || value.IndexOf('\\') >= 0 || value.StartsWith("~", StringComparison.Ordinal);
        }

        private static bool IsWithin(string value, string parent)
        {
            var normalizedValue = (value ?? string.Empty).TrimEnd('/');
            var normalizedParent = (parent ?? string.Empty).TrimEnd('/');
            return string.Equals(normalizedValue, normalizedParent, StringComparison.OrdinalIgnoreCase)
                || normalizedValue.StartsWith(normalizedParent + "/", StringComparison.OrdinalIgnoreCase);
        }

        private static bool IsReviewedAssetPath(string relativePath)
        {
            return relativePath.Equals("SiteAssets", StringComparison.OrdinalIgnoreCase)
                || relativePath.StartsWith("SiteAssets/", StringComparison.OrdinalIgnoreCase)
                || relativePath.Equals("Style Library", StringComparison.OrdinalIgnoreCase)
                || relativePath.StartsWith("Style Library/", StringComparison.OrdinalIgnoreCase);
        }

        private static string NormalizeTargetPath(string value, string targetWebPath)
        {
            var normalized = Uri.UnescapeDataString(value).Replace('\\', '/');
            if (!normalized.StartsWith("/", StringComparison.Ordinal))
            {
                normalized = targetWebPath.TrimEnd('/') + "/" + normalized.TrimStart('/');
            }
            if (!normalized.StartsWith(targetWebPath.TrimEnd('/') + "/", StringComparison.OrdinalIgnoreCase))
            {
                throw new ArgumentException("Target List path is outside its mapped target Web.", nameof(value));
            }
            return normalized.TrimEnd('/');
        }

        private static int IndexOf(IList<Guid> values, Guid value)
        {
            var index = values.IndexOf(value);
            return index < 0 ? int.MaxValue : index;
        }

        private static MigrationIssue Issue(string code, string subject, string message)
        {
            return new MigrationIssue
            {
                Code = code,
                Severity = MigrationIssueSeverity.Blocker,
                Subject = subject,
                Ingredient = "ListDependency",
                Message = message
            };
        }
    }
}
