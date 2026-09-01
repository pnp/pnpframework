using PnP.Framework.Migration.Diagnostics;
using PnP.Framework.Migration.Evidence;
using PnP.Framework.Migration.Lists.Capture;
using PnP.Framework.Migration.Lists.Fields;
using PnP.Framework.Migration.Lists.Items;
using PnP.Framework.Migration.Packaging;
using PnP.Framework.Migration.Schema.Fields;
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

        private static readonly string[] KnownRuntimeContentTypeIds =
        {
            "0x01", "0x0101", "0x0120", "0x0120D520", "0x0120D520A8", "0x0120D520A808"
        };

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
                plans.Add(CreateListPlan(source, owner, taxonomyMappings, targetOverride));
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
            plan.PlanDigest = null;
            plan.TargetProbe = null;
            try
            {
                return MigrationDigest.ComputeSha256(MigrationContractSerializer.SerializeCanonical(plan));
            }
            finally
            {
                plan.PlanDigest = value;
                plan.TargetProbe = probe;
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
                if (attachment.Content == null || attachment.Content.Availability != EvidenceAvailability.Captured || attachment.Content.Artifact == null)
                {
                    issues.Add(Issue("ListBinaryEvidenceUnavailable", "attachment:" + attachment.ServerRelativeUrl, "Exact attachment bytes are required before materialization."));
                }
            }
            foreach (var document in source.Items.Where(value => value.Document != null && value.Document.Kind == ListDocumentObjectKind.File).Select(value => value.Document))
            {
                if (document.Content == null || document.Content.Availability != EvidenceAvailability.Captured || document.Content.Artifact == null)
                {
                    issues.Add(Issue("ListBinaryEvidenceUnavailable", "document:" + document.ServerRelativeUrl, "Exact document bytes are required before materialization."));
                }
            }

            foreach (var contentType in source.ContentTypes)
            {
                var runtime = KnownRuntimeContentTypeIds.Any(value => string.Equals(contentType.ParentId, value, StringComparison.OrdinalIgnoreCase)
                    || (contentType.ParentId ?? string.Empty).StartsWith(value + "00", StringComparison.OrdinalIgnoreCase));
                if (!runtime)
                {
                    issues.Add(Issue("CustomListContentTypeClosureUnavailable", "content-type:" + contentType.Id,
                        "The List uses custom content type '" + contentType.Name + "'. Its site-content-type parent closure must be materialized before this List can execute."));
                }
            }

            var fieldPlans = source.Fields.Select(field => CreateFieldPlan(source, field, taxonomyMappings, issues))
                .OrderBy(value => value.Disposition == ListFieldMaterializationDisposition.CreateOrReuseOwnedCalculated ? 1 : 0)
                .ThenBy(value => value.InternalName, StringComparer.OrdinalIgnoreCase).ToList();
            var viewPlans = source.Views.Select(view => new ListViewMaterializationPlan
            {
                SourceViewId = view.Id,
                Title = view.Title,
                SourceServerRelativeUrl = view.ServerRelativeUrl,
                Disposition = view.Availability == EvidenceAvailability.Unavailable || view.Availability == EvidenceAvailability.Conflict
                    ? ListViewMaterializationDisposition.Block
                    : view.PersonalView
                        ? ListViewMaterializationDisposition.SkipPersonal
                        : view.IsPageBound
                            ? ListViewMaterializationDisposition.CreateOrReuseWebPartView
                            : ListViewMaterializationDisposition.CreateOrReuseOwnedPublicView,
                Source = view,
                Reason = view.PersonalView
                    ? "Personal views are user-scoped and remain evidence-only."
                    : view.IsPageBound
                        ? "The approved Web Part rewrite replays this embedded view with a target runtime view ID."
                        : "Create or exactly reuse the shared view inside the migration-owned target List."
            }).ToList();
            foreach (var blockedView in viewPlans.Where(value => value.Disposition == ListViewMaterializationDisposition.Block))
            {
                issues.Add(Issue("ViewEvidenceUnavailable", "view:" + blockedView.SourceViewId.ToString("D"), "View schema evidence is unavailable or conflicting."));
            }
            foreach (var customRenderingView in source.Views.Where(value => IsCustomRenderingReference(value.JsLink) || IsCustomRenderingReference(value.XslLink)))
            {
                issues.Add(Issue("ViewRenderingResourceUnavailable", "view:" + customRenderingView.Id.ToString("D"),
                    "Custom JSLink/XslLink requires separately captured exact rendering-resource bytes before the view can execute."));
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
                TargetWebServerRelativeUrl = owner.TargetServerRelativeUrl,
                TargetRootFolderServerRelativeUrl = targetPath,
                TargetTitle = string.IsNullOrWhiteSpace(targetOverride == null ? null : targetOverride.TargetTitle) ? source.Title : targetOverride.TargetTitle,
                OriginalIdentifier = "urn:pnp:spo-list:v1:" + source.SourceSiteId.ToString("D") + ":" + source.SourceWebId.ToString("D") + ":" + source.SourceListId.ToString("D"),
                Disposition = issues.Any(value => value.Severity == MigrationIssueSeverity.Blocker) ? ListMaterializationDisposition.Block : ListMaterializationDisposition.CreateOwned,
                Fields = fieldPlans,
                Views = viewPlans,
                Issues = issues.OrderBy(value => value.Code, StringComparer.Ordinal).ThenBy(value => value.Subject, StringComparer.Ordinal).ToList()
            };
            plan.PlanDigest = ComputePlanDigest(plan);
            return plan;
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
            var targetRuntime = field.FromBaseType || HasPlatformSource(field.SchemaXml);
            if (targetRuntime)
            {
                return Plan(field,
                    SupportedScalarTypes.Contains(field.TypeAsString)
                        ? ListFieldMaterializationDisposition.RequireTargetRuntimeAndCopyValue
                        : ListFieldMaterializationDisposition.RequireTargetRuntime,
                    null,
                    "The target List template owns this runtime field identity.");
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
                    "Create or reuse the taxonomy field with the approved target store/set; target SharePoint allocates WssIds.");
            }
            var hasValue = HasBusinessValue(source, field.InternalName);
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
            var requiredBySurface = source.Views.Any(view => view.ViewFields.Contains(field.InternalName, StringComparer.OrdinalIgnoreCase))
                || source.ContentTypes.Any(contentType => contentType.FieldLinks.Any(link => link.FieldId == field.Id));
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

        private static bool IsCustomRenderingReference(string value)
        {
            if (string.IsNullOrWhiteSpace(value))
            {
                return false;
            }
            return value.IndexOf('/') >= 0 || value.IndexOf('\\') >= 0 || value.StartsWith("~", StringComparison.Ordinal);
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
