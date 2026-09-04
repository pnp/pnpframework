using PnP.Framework.Migration.Lists.Planning;
using PnP.Framework.Migration.Pages.Publishing.Capture;
using PnP.Framework.Migration.Pages.Publishing.Planning;
using System;
using System.Collections.Generic;
using System.Linq;
using static PnP.Framework.Migration.Pages.Publishing.Reporting.Sections.MigrationReportSectionFormatter;

namespace PnP.Framework.Migration.Pages.Publishing.Reporting.Sections
{
    internal static class ListDependencyMigrationReportSection
    {
        public static void Append(MarkdownReportWriter writer, PublishingPageCaptureBundle snapshot, PublishingPageMigrationPlan plan)
        {
            writer.Heading(2, $"List dependency closure ({snapshot.ListDependencies.Count})");
            writer.Paragraph("Every returned source item field value is retained as typed evidence plus best-effort raw evidence. Import writes only field dispositions explicitly understood by the current materializer. Lookup Lists are ordered before their consumers; source List, View, item and WssId identities are never forced onto the target.");
            writer.Table("Lookup dependency edges", new[] { "Consumer List", "Lookup List", "Field ID", "Field internal name", "How to read it" },
                snapshot.ListLookupDependencies.Select(value => Row(
                    value.SourceListId,
                    value.LookupListId,
                    value.FieldId,
                    value.FieldInternalName,
                    "The lookup List and its source-to-target item-ID catalog must exist before this field value is written.")));

            var planSet = plan.ListMigration;
            writer.Table("List dependency execution order", new[] { "Order", "Source List ID", "Plan present" },
                (planSet == null ? Enumerable.Empty<Guid>() : planSet.OrderedSourceListIds)
                    .Select((value, index) => Row(index + 1, value, planSet.Lists.Any(item => item.SourceListId == value))));
            var plans = planSet == null
                ? new Dictionary<Guid, ListMaterializationPlan>()
                : planSet.Lists.ToDictionary(value => value.SourceListId);

            foreach (var source in snapshot.ListDependencies.OrderBy(value => value.SourceWebUrl, StringComparer.OrdinalIgnoreCase)
                         .ThenBy(value => value.RootFolderServerRelativeUrl, StringComparer.OrdinalIgnoreCase))
            {
                ListMaterializationPlan listPlan;
                plans.TryGetValue(source.SourceListId, out listPlan);
                writer.Heading(3, "List: " + PublishingPageReportValueFormatter.EscapeHeading(source.Title) + " (" + source.SourceListId.ToString("D") + ")");
                writer.Table(null, new[] { "JSON property", "Captured value", "How to read it" }, new[]
                {
                    Row("schemaVersion", source.SchemaVersion, "Nested List dependency contract version."),
                    Row("sourceSiteId", source.SourceSiteId, "Source SPSite identity evidence."),
                    Row("sourceWebId", source.SourceWebId, "Source owning SPWeb identity used by the topology map."),
                    Row("sourceWebUrl", source.SourceWebUrl, "Source owning Web URL."),
                    Row("sourceListId", source.SourceListId, "Source List identity evidence; target creates a new List ID."),
                    Row("title", source.Title, "Source display title; an explicit target override may change only the target title."),
                    Row("description", source.Description, "Description applied to a newly created target List."),
                    Row("templateFeatureId", source.TemplateFeatureId, "Feature identity supplied with List creation when non-empty."),
                    Row("baseTemplate", source.BaseTemplate, "Numeric List template; only reviewed templates are executable."),
                    Row("baseType", source.BaseType, "GenericList or DocumentLibrary storage family."),
                    Row("rootFolderServerRelativeUrl", source.RootFolderServerRelativeUrl, "Web-local source path mapped under the target owner Web."),
                    Row("hidden", source.Hidden, "Target List visibility setting."),
                    Row("contentTypesEnabled", source.ContentTypesEnabled, "Whether List content types are enabled."),
                    Row("enableAttachments", source.EnableAttachments, "Whether list-item attachments are allowed."),
                    Row("enableFolderCreation", source.EnableFolderCreation, "Whether folder creation is enabled."),
                    Row("enableVersioning", source.EnableVersioning, "Current List versioning setting; history itself is not copied."),
                    Row("enableMinorVersions", source.EnableMinorVersions, "Current minor-version setting; history itself is not copied."),
                    Row("enableModeration", source.EnableModeration, "Current moderation setting."),
                    Row("forceCheckout", source.ForceCheckout, "Current required-checkout setting."),
                    Row("informationRightsManagement.irmEnabled", source.InformationRightsManagement?.IrmEnabled, "Whether the source library dynamically applies Information Rights Management."),
                    Row("informationRightsManagement.irmExpire", source.InformationRightsManagement?.IrmExpire, "Whether source IRM protection expires."),
                    Row("informationRightsManagement.irmReject", source.InformationRightsManagement?.IrmReject, "Whether the source library rejects documents that do not support IRM."),
                    Row("informationRightsManagement.availability", source.InformationRightsManagement?.Availability, "Captured means the library IRM state and, when enabled, its detailed policy were read."),
                    Row("informationRightsManagement.policyTitle", source.InformationRightsManagement?.Policy?.PolicyTitle, "Source IRM policy title."),
                    Row("informationRightsManagement.templateId", source.InformationRightsManagement?.Policy?.TemplateId, "Source rights-management template identity, when configured."),
                    Row("sourceItemCount", source.SourceItemCount, "Must equal the number of captured item snapshots for an executable plan."),
                    Row("fields.count", source.Fields.Count, "Every captured List field definition."),
                    Row("contentTypes.count", source.ContentTypes.Count, "Every captured List-local content type."),
                    Row("hasExplicitUniqueContentTypeOrder", source.HasExplicitUniqueContentTypeOrder, "False means SharePoint returned null, which exposes all allowed content types without an explicit order."),
                    Row("uniqueContentTypeOrder", Join(source.UniqueContentTypeOrder), "Source List-local IDs in New-button order; mapped to target-generated List content type IDs. Platform Folder/UntypedDocument children are filtered because SharePoint rejects them in this property."),
                    Row("siteContentTypes.count", source.SiteContentTypes.Count, "Custom site-content-type ancestor closure required by the List."),
                    Row("views.count", source.Views.Count, "Public, page-bound, and personal View evidence; personal Views are not restored."),
                    Row("items.count", source.Items.Count, "Current folders/items/files captured; version history is not included."),
                    Row("availability", source.Availability, "Partial/unavailable evidence blocks exact materialization where required."),
                    Row("diagnostics", Join(source.Diagnostics), "List-level acquisition findings.")
                });
                AppendPlan(writer, listPlan);
                ListSchemaMigrationReportSection.Append(writer, source, listPlan);
                ListItemMigrationReportSection.Append(writer, source);
            }

            if (planSet == null)
            {
                writer.Paragraph("No List migration plan set was produced.");
            }
            else
            {
                writer.Table("List plan-set identity", new[] { "Property", "Value", "How to read it" }, new[]
                {
                    Row("schemaVersion", planSet.SchemaVersion, "Nested List-plan contract version."),
                    Row("planDigest", planSet.PlanDigest, "Digest over dependency order, per-List plans, probes, admissions, and issues."),
                    Row("isExecutable", planSet.IsExecutable, "True only when every List, field, view, schema closure, lookup edge, and target probe is admitted.")
                });
                AppendIssues(writer, "List plan-set issues", planSet.Issues);
            }
        }

        private static void AppendPlan(MarkdownReportWriter writer, ListMaterializationPlan plan)
        {
            writer.Heading(4, "Target List plan and probe");
            if (plan == null)
            {
                writer.Paragraph("No target plan covers this captured List.");
                return;
            }
            writer.Table(null, new[] { "Property", "Value", "How to read it" }, new[]
            {
                Row("targetWebUrl", plan.TargetWebUrl, "Mapped target owner Web."),
                Row("targetWebServerRelativeUrl", plan.TargetWebServerRelativeUrl, "Target Web path boundary."),
                Row("preferredTargetRootFolderServerRelativeUrl", plan.PreferredTargetRootFolderServerRelativeUrl, "Exact relative-path List/library root before live collision resolution."),
                Row("targetRootFolderServerRelativeUrl", plan.TargetRootFolderServerRelativeUrl, "Final approved List/library root; only the colliding List leaf may differ."),
                Row("preferredTargetTitle", plan.PreferredTargetTitle, "Source-equivalent target title before collision resolution."),
                Row("targetTitle", plan.TargetTitle, "Final approved title; it changes only when SharePoint title uniqueness also collides."),
                Row("originalIdentifier", plan.OriginalIdentifier, "Source-qualified provenance written to the target List root folder."),
                Row("collisionResolved", plan.TargetProbe?.CollisionResolved, "True when planning allocated a stable suffix instead of treating a foreign target object as terminal."),
                Row("collisionResolutionReason", plan.TargetProbe?.CollisionResolutionReason, "Retained evidence for the path/title decision."),
                Row("disposition", plan.Disposition, "CreateOwned creates; ReuseOwned requires exact provenance and semantic digest; local Block records a capability gap and becomes final ingredient Defer."),
                Row("planDigest", plan.PlanDigest, "Semantic identity written with provenance and used for resume."),
                Row("isExecutable", plan.IsExecutable, "Includes List issues, custom site-content-type closure, and target admission.")
            });
            var probe = plan.TargetProbe;
            if (probe != null)
            {
                writer.Table(null, new[] { "Target probe property", "Observed value", "Interpretation" }, new[]
                {
                    Row("targetWebUrl", probe.TargetWebUrl, "Web actually probed or planned for topology-first creation."),
                    Row("preferredTargetRootFolderServerRelativeUrl", probe.PreferredTargetRootFolderServerRelativeUrl, "Exact mapped path before collision resolution."),
                    Row("targetRootFolderServerRelativeUrl", probe.TargetRootFolderServerRelativeUrl, "Final sealed List/library path."),
                    Row("preferredTargetTitle", probe.PreferredTargetTitle, "Mapped target title before collision resolution."),
                    Row("targetTitle", probe.TargetTitle, "Final sealed target title."),
                    Row("collisionResolved", probe.CollisionResolved, "Whether planning moved only the colliding List node."),
                    Row("collisionResolutionReason", probe.CollisionResolutionReason, "Why the preferred target could not be used."),
                    Row("targetWebExists", probe.TargetWebExists, "False is allowed only when explicitly deferred until topology materialization."),
                    Row("deferredUntilTopologyMaterialization", probe.DeferredUntilTopologyMaterialization, "Import must create/recover the mapped owner Web before re-running this probe."),
                    Row("targetWebId", probe.TargetWebId, "Observed runtime Web identity."),
                    Row("listExists", probe.ListExists, "Whether the approved target path is occupied."),
                    Row("targetListId", probe.TargetListId, "Observed runtime List identity when present."),
                    Row("existingTitle", probe.ExistingTitle, "Observed title at the approved path."),
                    Row("existingBaseTemplate", probe.ExistingBaseTemplate, "Observed template at the approved path."),
                    Row("existingOriginalIdentifier", probe.ExistingOriginalIdentifier, "Observed migration ownership marker."),
                    Row("existingPlanDigest", probe.ExistingPlanDigest, "Observed semantic ownership digest."),
                    Row("sameTitleDifferentPaths", Join(probe.SameTitleDifferentPaths), "Existing Lists that would trigger SharePoint's title uniqueness collision even when the path is free."),
                    Row("canManageLists", probe.CanManageLists, "Caller had ManageLists during planning."),
                    Row("disposition", probe.Disposition, "Sealed create/reuse/block result."),
                    Row("isAdmitted", probe.IsAdmitted, "No blocking target issue was observed.")
                });
                AppendIssues(writer, null, probe.Issues, 4);
            }
            AppendIssues(writer, null, plan.Issues, 4);
        }
    }
}
