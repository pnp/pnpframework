using PnP.Framework.Migration.Lists.Capture;
using PnP.Framework.Migration.Lists.Planning;
using PnP.Framework.Migration.Features;
using PnP.Framework.Migration.Schema.ContentTypes;
using PnP.Framework.Migration.Schema.Fields;
using System;
using System.Collections.Generic;
using System.Linq;
using static PnP.Framework.Migration.Pages.Publishing.Reporting.Sections.MigrationReportSectionFormatter;

namespace PnP.Framework.Migration.Pages.Publishing.Reporting.Sections
{
    internal static class ListSchemaMigrationReportSection
    {
        public static void Append(MarkdownReportWriter writer, ListDependencySnapshot source, ListMaterializationPlan plan)
        {
            AppendFields(writer, source, plan);
            AppendPlatformFeatures(writer, plan);
            AppendSiteContentTypes(writer, source, plan);
            AppendListContentTypes(writer, source);
            AppendViews(writer, source, plan);
        }

        private static void AppendPlatformFeatures(MarkdownReportWriter writer, ListMaterializationPlan plan)
        {
            var features = plan == null ? Array.Empty<PlatformFeatureMaterializationPlan>() : plan.RequiredFeatures.ToArray();
            writer.Heading(4, $"Required SharePoint platform features ({features.Length})");
            writer.Paragraph("Conditional target-runtime content types are not assumed to exist merely because SharePoint owns their IDs. The plan probes their providing site feature and either reuses or explicitly activates that platform capability before List schema is materialized.");
            writer.Table(null,
                new[] { "Feature ID / name", "Scope / order", "Required by content types", "Expected runtime content types", "Target probe", "Plan disposition", "Reason / diagnostics" },
                features.Select(feature => Row(
                    $"{feature.FeatureId:D} / {Format(feature.Name)}",
                    $"{feature.Scope} / {feature.DependencyOrder}",
                    Join(feature.RequiredByContentTypeIds),
                    Join(feature.ExpectedContentTypeIds),
                    feature.TargetProbe == null
                        ? null
                        : $"active={feature.TargetProbe.IsActive}; canActivate={feature.TargetProbe.CanActivate}; deferred={feature.TargetProbe.DeferredUntilTopologyMaterialization}; availableCTs={Join(feature.TargetProbe.AvailableContentTypeIds)}",
                    feature.Disposition,
                    Join(new[] { feature.Reason }.Concat(feature.TargetProbe == null
                        ? Enumerable.Empty<string>()
                        : feature.TargetProbe.Issues.Select(value => value.Message))))));
        }

        private static void AppendFields(MarkdownReportWriter writer, ListDependencySnapshot source, ListMaterializationPlan plan)
        {
            writer.Heading(4, $"List fields ({source.Fields.Count})");
            var actions = plan == null
                ? new Dictionary<Guid, ListFieldMaterializationPlan>()
                : plan.Fields.ToDictionary(value => value.SourceFieldId);
            writer.Table(null,
                new[] { "Field ID", "Internal / display name", "Type / group", "Flags", "Lookup / taxonomy binding", "Source schema", "Availability", "Plan disposition", "Target schema", "Reason / diagnostics" },
                source.Fields.OrderBy(value => value.InternalName, StringComparer.OrdinalIgnoreCase).Select(field =>
                {
                    ListFieldMaterializationPlan action;
                    actions.TryGetValue(field.Id, out action);
                    return Row(
                        field.Id,
                        $"internal={Format(field.InternalName)}; title={Format(field.Title)}",
                        $"type={Format(field.TypeAsString)}; group={Format(field.Group)}",
                        $"hidden={field.Hidden}; readOnly={field.ReadOnly}; required={field.Required}; fromBaseType={field.FromBaseType}; sealed={field.Sealed}",
                        $"lookupWeb={Format(field.SourceLookupWebId)}; lookupList={Format(field.SourceLookupListId)}; lookupField={Format(field.LookupField)}; taxonomy={FormatTaxonomy(field.Taxonomy)}",
                        $"xml={Summarize(field.SchemaXml)}; sha256={Format(field.SchemaXmlSha256)}; portableSha256={Format(field.PortableSchemaSha256)}",
                        field.Availability,
                        action == null ? null : $"{action.Disposition}; plannedTitle={Format(action.Title)}",
                        action == null ? null : $"xml={Summarize(action.TargetSchemaXml)}; portableSha256={Format(action.TargetPortableSchemaSha256)}",
                        Join(new[] { action?.Reason }.Concat(field.Diagnostics)));
                }));
        }

        private static void AppendSiteContentTypes(MarkdownReportWriter writer, ListDependencySnapshot source, ListMaterializationPlan plan)
        {
            writer.Heading(4, $"Custom site content-type closure ({source.SiteContentTypes.Count})");
            writer.Paragraph("Only non-runtime ancestors required by List-local content types are captured. Exact runtime IDs such as Item or Document terminate the closure; a custom child of Document is still custom and is not classified by prefix alone.");
            var nodes = plan == null
                ? new Dictionary<string, ContentTypeClosureNodePlan>(StringComparer.OrdinalIgnoreCase)
                : plan.SiteContentTypes.ToDictionary(value => value.Schema.ContentTypeId, StringComparer.OrdinalIgnoreCase);
            foreach (var contentType in source.SiteContentTypes.OrderBy(value => value.ContentTypeId == null ? 0 : value.ContentTypeId.Length)
                         .ThenBy(value => value.ContentTypeId, StringComparer.OrdinalIgnoreCase))
            {
                ContentTypeClosureNodePlan node;
                nodes.TryGetValue(contentType.ContentTypeId, out node);
                writer.Heading(5, PublishingPageReportValueFormatter.EscapeHeading(contentType.Name) + " (" + Format(contentType.ContentTypeId) + ")");
                writer.Table(null, new[] { "Property", "Source evidence", "Target plan / interpretation" }, new[]
                {
                    Row("schemaVersion", contentType.SchemaVersion, "Nested content-type schema contract."),
                    Row("evidenceState / availability", $"{contentType.EvidenceState} / {contentType.Availability}", "Readable/Captured is required for creation."),
                    Row("sourceWebUrl / sourceScope", $"{Format(contentType.SourceWebUrl)} / {Format(contentType.SourceScope)}", node == null ? null : $"owner={node.SourceOwnerWebId:D}; targetOwner={Format(node.TargetOwnerWebUrl)}"),
                    Row("contentTypeId", contentType.ContentTypeId, node?.Schema?.Disposition),
                    Row("name / description / group", $"{Format(contentType.Name)} / {Format(contentType.Description)} / {Format(contentType.Group)}", node?.Schema?.Reason),
                    Row("readOnly / sealed / hidden", $"{contentType.ReadOnly} / {contentType.Sealed} / {contentType.Hidden}", "Exact flags carried into target admission, creation, and readback."),
                    Row("parentContentTypeId / name", $"{Format(contentType.ParentContentTypeId)} / {Format(contentType.ParentContentTypeName)}", "Parent is materialized first unless it is a target-runtime content type."),
                    Row("requiredFieldLinks.count", contentType.RequiredFieldLinks.Count, "Minimal direct/inherited field-link closure."),
                    Row("requiredFieldClosure.count", contentType.RequiredFieldClosure.Count, "Complete schema evidence for every required field."),
                    Row("sources", Join(contentType.Sources.Select(FormatEvidenceSource)), "Evidence lineage."),
                    Row("diagnostics", Join(contentType.Diagnostics), "Capture findings."),
                    Row("node.planDigest", node?.PlanDigest, "Semantic node identity excluding mutable target probe/admission."),
                    Row("node.deferredUntilTopologyMaterialization", node?.DeferredUntilTopologyMaterialization, "Target probe is deliberately postponed when its owner Web will be created first."),
                    Row("node.targetAdmission", node?.TargetAdmission == null ? null : $"eligible={node.TargetAdmission.IsEligible}; disposition={node.TargetAdmission.Disposition}", "Freshly re-evaluated during import.")
                });
                writer.Table(null, new[] { "Field link ID", "Name", "Required", "Hidden", "Role" },
                    contentType.RequiredFieldLinks.Select(value => Row(value.FieldId, value.Name, value.Required, value.Hidden, value.Role)));
                var fieldPlans = node?.Schema?.Fields == null
                    ? new Dictionary<Guid, FieldSchemaMaterializationPlan>()
                    : node.Schema.Fields.ToDictionary(value => value.FieldId);
                writer.Table(null,
                    new[] { "Field ID", "Internal / title", "Type / group", "Flags", "Role / ownership", "Taxonomy binding", "Source schema", "Plan disposition / target mapping", "Reason / sources / diagnostics" },
                    contentType.RequiredFieldClosure.OrderBy(value => value.InternalName, StringComparer.OrdinalIgnoreCase).Select(field =>
                    {
                        FieldSchemaMaterializationPlan fieldPlan;
                        fieldPlans.TryGetValue(field.Id, out fieldPlan);
                        return Row(
                            field.Id,
                            $"internal={Format(field.InternalName)}; title={Format(field.Title)}",
                            $"type={Format(field.TypeAsString)}; group={Format(field.Group)}",
                            $"required={field.Required}; hidden={field.Hidden}; readOnly={field.ReadOnly}; sealed={field.Sealed}",
                            $"role={field.Role}; ownership={field.Ownership}",
                            FormatTaxonomy(field.Taxonomy),
                            $"xml={Summarize(field.SchemaXml)}; sha256={Format(field.SchemaXmlSha256)}; portableSha256={Format(field.PortableSchemaSha256)}",
                            fieldPlan == null ? null : $"{fieldPlan.Disposition}; targetPortableSha256={Format(fieldPlan.TargetPortableSchemaSha256)}; targetStore={Format(fieldPlan.TargetTermStoreId)}; targetSet={Format(fieldPlan.TargetTermSetId)}",
                            Join(new[] { fieldPlan?.Reason }.Concat(field.Sources.Select(FormatEvidenceSource)).Concat(field.Diagnostics)));
                    }));
                if (node?.TargetProbe != null)
                {
                    writer.Table(null, new[] { "Target content-type probe", "Observed value" }, new[]
                    {
                        Row("contentTypeId", node.TargetProbe.ContentTypeId),
                        Row("parentContentTypeAvailable / resolvedParent", $"{node.TargetProbe.ParentContentTypeAvailable} / {Format(node.TargetProbe.ResolvedParentContentTypeId)}"),
                        Row("contentTypeExists", node.TargetProbe.ContentTypeExists),
                        Row("existingName / description / group", $"{Format(node.TargetProbe.ExistingName)} / {Format(node.TargetProbe.ExistingDescription)} / {Format(node.TargetProbe.ExistingGroup)}"),
                        Row("existingReadOnly / sealed / hidden", $"{node.TargetProbe.ExistingReadOnly} / {node.TargetProbe.ExistingSealed} / {node.TargetProbe.ExistingHidden}"),
                        Row("existingParentContentTypeId", node.TargetProbe.ExistingParentContentTypeId),
                        Row("sameNameDifferentIds", Join(node.TargetProbe.SameNameDifferentIds)),
                        Row("canManageContentTypes", node.TargetProbe.CanManageContentTypes),
                        Row("availability", node.TargetProbe.Availability),
                        Row("diagnostics", Join(node.TargetProbe.Diagnostics))
                    });
                }
                if (node?.TargetAdmission != null)
                {
                    AppendIssues(writer, null, node.TargetAdmission.Issues, 5);
                    if (node.TargetAdmission.Warnings.Count > 0)
                    {
                        writer.Table(null, new[] { "Target-admission warning" }, node.TargetAdmission.Warnings.Select(value => Row(value)));
                    }
                }
            }
        }

        private static void AppendListContentTypes(MarkdownReportWriter writer, ListDependencySnapshot source)
        {
            writer.Heading(4, $"List-local content types ({source.ContentTypes.Count})");
            foreach (var contentType in source.ContentTypes.OrderBy(value => value.Id, StringComparer.OrdinalIgnoreCase))
            {
                writer.Heading(5, PublishingPageReportValueFormatter.EscapeHeading(contentType.Name) + " (" + Format(contentType.Id) + ")");
                writer.Table(null, new[] { "Property", "Value", "How to read it" }, new[]
                {
                    Row("id", contentType.Id, "Source List-local ID; target List generates a new child ID recorded in the receipt."),
                    Row("parentId", contentType.ParentId, "Exact site content type added to the target List."),
                    Row("name", contentType.Name, "Replayed after the exact site parent resolves the target List-local content type; matching never relies on a potentially customized name."),
                    Row("description", contentType.Description, "Replayed and freshly verified metadata."),
                    Row("group", contentType.Group, "Replayed and freshly verified metadata."),
                    Row("hidden / readOnly / sealed", $"{contentType.Hidden} / {contentType.ReadOnly} / {contentType.Sealed}", "Replayed and freshly verified source flags."),
                    Row("fieldLinks.count", contentType.FieldLinks.Count, "Every captured List-local field-link setting.")
                });
                writer.Table(null, new[] { "Field ID", "Internal / display name", "Required", "Hidden", "ReadOnly" },
                    contentType.FieldLinks.OrderBy(value => value.InternalName, StringComparer.OrdinalIgnoreCase)
                        .Select(value => Row(value.FieldId, $"{Format(value.InternalName)} / {Format(value.DisplayName)}", value.Required, value.Hidden, value.ReadOnly)));
            }
        }

        private static void AppendViews(MarkdownReportWriter writer, ListDependencySnapshot source, ListMaterializationPlan plan)
        {
            writer.Heading(4, $"Views ({source.Views.Count})");
            var actions = plan == null
                ? new Dictionary<Guid, ListViewMaterializationPlan>()
                : plan.Views.ToDictionary(value => value.SourceViewId);
            writer.Table(null,
                new[] { "View ID", "Title / URL", "Kind / flags", "Query / fields", "Rendering", "Captured XML", "Availability", "Plan disposition", "Reason / diagnostics" },
                source.Views.OrderBy(value => value.Title, StringComparer.OrdinalIgnoreCase).ThenBy(value => value.Id).Select(view =>
                {
                    ListViewMaterializationPlan action;
                    actions.TryGetValue(view.Id, out action);
                    return Row(
                        view.Id,
                        $"title={Format(view.Title)}; url={Format(view.ServerRelativeUrl)}",
                        $"type={Format(view.ViewType)}; hidden={view.Hidden}; default={view.DefaultView}; personal={view.PersonalView}; pageBound={view.IsPageBound}; rowLimit={view.RowLimit}; paged={view.Paged}",
                        $"query={Summarize(view.ViewQuery)}; fields={Join(view.ViewFields)}",
                        $"jsLink={Format(view.JsLink)}; xslLink={Format(view.XslLink)}",
                        $"sha256={Format(view.ListViewXmlSha256)}; xml={Summarize(view.ListViewXml)}",
                        view.Availability,
                        action?.Disposition,
                        Join(new[] { action?.Reason }.Concat(view.Diagnostics)));
                }));

            writer.Heading(5, $"Custom rendering resources ({source.ViewRenderingResources.Count})");
            var resourcePlans = plan == null
                ? new Dictionary<string, ListViewRenderingResourceMaterializationPlan>(StringComparer.Ordinal)
                : plan.ViewRenderingResources.ToDictionary(value => value.SourceResourceId, StringComparer.Ordinal);
            writer.Table(null,
                new[] { "Resource ID", "Kind", "Source", "Artifact", "Availability", "Target / disposition", "Consumers", "Reason / diagnostics" },
                source.ViewRenderingResources.OrderBy(value => value.SourceServerRelativeUrl, StringComparer.OrdinalIgnoreCase).Select(resource =>
                {
                    resourcePlans.TryGetValue(resource.Id, out var resourcePlan);
                    var consumers = source.Views
                        .Where(view => view.RenderingResourceBindings.Any(binding => string.Equals(binding.ResourceId, resource.Id, StringComparison.Ordinal)))
                        .Select(view => view.Id.ToString("D"));
                    return Row(
                        resource.Id,
                        resource.Kind,
                        $"url={Format(resource.SourceAbsoluteUrl)}; path={Format(resource.SourceServerRelativeUrl)}",
                        $"sha256={Format(resource.Artifact?.Sha256)}; length={resource.Artifact?.Length}",
                        resource.Availability,
                        $"path={Format(resourcePlan?.TargetServerRelativeUrl)}; disposition={resourcePlan?.Disposition}",
                        Join(consumers),
                        Join(new[] { resourcePlan?.Reason }.Concat(resource.Diagnostics)));
                }));
        }
    }
}
