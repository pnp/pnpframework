using PnP.Framework.Migration.Pages.Publishing.Capture;
using PnP.Framework.Migration.Pages.Publishing.Planning;
using PnP.Framework.Migration.Topology;
using System;
using System.Linq;
using static PnP.Framework.Migration.Pages.Publishing.Reporting.Sections.MigrationReportSectionFormatter;

namespace PnP.Framework.Migration.Pages.Publishing.Reporting.Sections
{
    internal static class TopologyMigrationReportSection
    {
        public static void Append(MarkdownReportWriter writer, PublishingPageCaptureBundle snapshot, PublishingPageMigrationPlan plan)
        {
            writer.Heading(2, "Site collection and Web topology");
            writer.Paragraph("Topology preserves SharePoint ownership level: a source site collection maps to a target site collection, each source Web maps to one target Web under the mapped parent, and Web-owned Lists/files/pages remain under that mapped Web. Source GUIDs are evidence; target GUIDs are runtime-generated and recorded in the import receipt.");
            if (snapshot.SourceTopology == null)
            {
                writer.Paragraph("No source topology closure was required by this package.");
                return;
            }

            var source = snapshot.SourceTopology;
            writer.Table(null, new[] { "JSON path", "Value", "How to read it" }, new[]
            {
                Row("snapshot.sourceTopology.schemaVersion", source.SchemaVersion, "Nested source-topology contract version."),
                Row("snapshot.sourceTopology.siteId", source.SiteId, "Source SPSite identity; never forced onto the target."),
                Row("snapshot.sourceTopology.siteCollectionUrl", source.SiteCollectionUrl, "Source site-collection boundary."),
                Row("snapshot.sourceTopology.serverRelativeUrl", source.ServerRelativeUrl, "Source site-collection path boundary."),
                Row("snapshot.sourceTopology.rootWebId", source.RootWebId, "Source root SPWeb identity."),
                Row("snapshot.sourceTopology.availability", source.Availability, "Captured is required for executable topology planning."),
                Row("snapshot.sourceTopology.diagnostics", Join(source.Diagnostics), "Source acquisition findings.")
            });
            writer.Table("Source Web inventory", new[] { "Kind", "Site ID", "Web ID", "Parent Web ID", "URL / path", "Title", "Template", "Availability", "Diagnostics" },
                source.Webs.OrderBy(value => value.ServerRelativeUrl, StringComparer.OrdinalIgnoreCase).Select(value => Row(
                    value.ParentWebId.HasValue ? TopologyNodeKind.ChildWeb : TopologyNodeKind.SiteCollectionRoot,
                    value.SiteId,
                    value.WebId,
                    value.ParentWebId,
                    $"absolute={Format(value.WebUrl)}; serverRelative={Format(value.ServerRelativeUrl)}",
                    value.Title,
                    $"{Format(value.WebTemplate)}#{value.Configuration}",
                    value.Availability,
                    Join(value.Diagnostics))));

            if (plan.Topology == null)
            {
                writer.Paragraph("The target plan contains no topology mapping. If source topology exists, this normally means planning is blocked before a topology could be sealed.");
                return;
            }

            writer.Table("Target site-collection mapping", new[] { "Source Site ID", "Mode", "Preferred target", "Final target / expected ID", "Collision resolution", "Title / owner", "Template / locale", "Original identifier" },
                plan.Topology.SiteCollections.OrderBy(value => value.SourceSiteId).Select(value => Row(
                    value.SourceSiteId,
                    value.TargetMode,
                    value.PreferredTargetSiteCollectionUrl,
                    $"url={Format(value.TargetSiteCollectionUrl)}; expectedId={Format(value.ExpectedTargetSiteId)}",
                    $"resolved={value.TargetSiteCollisionResolved}; reason={Format(value.TargetSiteResolutionReason)}",
                    $"title={Format(value.TargetTitle)}; owner={Format(value.TargetOwner)}",
                    $"{Format(value.TargetTemplate)}; language={value.TargetLanguage}; timeZone={value.TargetTimeZone}",
                    value.OriginalIdentifier)));
            writer.Table("Approved Web mappings", new[] { "Kind", "Source Site / Web / parent", "Source URL / path", "Preferred target", "Final target / parent", "Target title / template", "Original identifier", "Mapping SHA-256" },
                plan.Topology.SiteCollections.SelectMany(value => value.Webs)
                    .OrderBy(value => value.TargetServerRelativeUrl, StringComparer.OrdinalIgnoreCase)
                    .Select(value => Row(
                        value.Kind,
                        $"site={value.SourceSiteId:D}; web={value.SourceWebId:D}; parent={Format(value.SourceParentWebId)}",
                        $"url={Format(value.SourceWebUrl)}; path={Format(value.SourceServerRelativeUrl)}",
                        $"url={Format(value.PreferredTargetWebUrl)}; path={Format(value.PreferredTargetServerRelativeUrl)}",
                        $"url={Format(value.TargetWebUrl)}; path={Format(value.TargetServerRelativeUrl)}; parent={Format(value.TargetParentWebUrl)}",
                        $"title={Format(value.TargetTitle)}; template={Format(value.TargetTemplate)}#{value.TargetConfiguration}",
                        value.OriginalIdentifier,
                        TopologyPlanner.ComputeWebMappingDigest(value))));
            writer.Table("Topology plan identity", new[] { "Property", "Value", "How to read it" }, new[]
            {
                Row("plan.topology.schemaVersion", plan.Topology.SchemaVersion, "Nested topology plan contract version."),
                Row("plan.topology.planDigest", plan.Topology.PlanDigest, "Semantic digest used for Web ownership claims and resume."),
                Row("computedPlanDigest", TopologyPlanner.ComputeDigest(plan.Topology), "Must equal plan.topology.planDigest.")
            });

            var analysis = plan.TopologyTargetAnalysis;
            if (analysis == null)
            {
                writer.Paragraph("No target topology analysis was sealed. An executable topology plan requires one.");
                return;
            }
            writer.Table("Target site-collection probes", new[] { "Source Site ID", "Preferred / final target", "Collision resolution", "Exists", "Target Site / root Web", "Disposition", "Admitted", "Issues" },
                analysis.SiteCollections.OrderBy(value => value.SourceSiteId).Select(value => Row(
                    value.SourceSiteId,
                    $"preferred={Format(value.PreferredTargetSiteCollectionUrl)}; final={Format(value.TargetSiteCollectionUrl)}",
                    $"resolved={value.CollisionResolved}; reason={Format(value.CollisionResolutionReason)}",
                    value.Exists,
                    $"site={Format(value.TargetSiteId)}; rootWeb={Format(value.TargetRootWebId)}",
                    value.Disposition,
                    value.IsAdmitted,
                    IssueSummary(value.Issues))));
            writer.Table("Target Web probes", new[] { "Source Site / Web", "Preferred / final target", "Collision resolution", "Exists / runtime identity", "Observed shape", "Observed provenance", "Disposition", "Admitted", "Issues" },
                analysis.SiteCollections.SelectMany(value => value.Webs)
                    .OrderBy(value => value.TargetWebUrl, StringComparer.OrdinalIgnoreCase)
                    .Select(value => Row(
                        $"site={value.SourceSiteId:D}; web={value.SourceWebId:D}",
                        $"preferred={Format(value.PreferredTargetWebUrl)} ({Format(value.PreferredTargetServerRelativeUrl)}); final={Format(value.TargetWebUrl)} ({Format(value.TargetServerRelativeUrl)})",
                        $"resolved={value.CollisionResolved}; reason={Format(value.CollisionResolutionReason)}",
                        $"exists={value.Exists}; site={Format(value.TargetSiteId)}; web={Format(value.TargetWebId)}; parent={Format(value.TargetParentWebId)}",
                        $"title={Format(value.ExistingTitle)}; template={Format(value.ExistingTemplate)}#{Format(value.ExistingConfiguration)}",
                        $"originalIdentifier={Format(value.ExistingOriginalIdentifier)}; planDigest={Format(value.ExistingPlanDigest)}",
                        value.Disposition,
                        value.IsAdmitted,
                        IssueSummary(value.Issues))));
            AppendIssues(writer, "Topology target-analysis issues", analysis.Issues);
        }
    }
}
