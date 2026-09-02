using Microsoft.SharePoint.Client;
using PnP.Framework.Migration.Lists.Planning;
using PnP.Framework.Migration.Pages.ClassicWebParts.Planning;
using PnP.Framework.Migration.Pages.Planning;
using PnP.Framework.Migration.Pages.Publishing.Capture;
using PnP.Framework.Migration.Topology;
using System;
using System.Collections.Generic;
using System.Linq;

namespace PnP.Framework.Migration.Pages.Publishing.Planning
{
    internal static class PublishingPageDependencyPlanner
    {
        public static PublishingPageDependencyPlan Build(
            ClientContext targetContext,
            PublishingPageCaptureBundle snapshot,
            Web targetWeb,
            Site targetSite,
            Web targetRootWeb,
            PagePlanningOptions options,
            ICollection<string> blockers,
            ICollection<string> warnings)
        {
            var result = new PublishingPageDependencyPlan();
            if (snapshot.SourceTopology != null)
            {
                var topologyResult = new TopologyPlanner().Build(
                    new[] { snapshot.SourceTopology },
                    new[]
                    {
                        new TargetSiteCollectionSpec
                        {
                            SourceSiteId = snapshot.SourceTopology.SiteId,
                            Mode = TargetSiteMode.ExistingTargetSite,
                            TargetSiteUrl = targetRootWeb.Url,
                            ExpectedTargetSiteId = targetSite.Id,
                            Title = targetRootWeb.Title,
                            Template = targetRootWeb.WebTemplate
                        }
                    },
                    options.TopologyPolicy);
                AddIssues(topologyResult.Issues, blockers);
                result.Topology = topologyResult.Plan;
                if (result.Topology != null)
                {
                    result.TopologyTargetAnalysis = TopologyTargetInspector.Inspect(targetContext, result.Topology, targetWeb.Url);
                    AddIssues(result.TopologyTargetAnalysis.Issues, blockers);
                    var pageWebMapping = result.Topology.SiteCollections.SelectMany(value => value.Webs)
                        .SingleOrDefault(value => value.SourceWebId == snapshot.Source.WebId);
                    if (pageWebMapping == null
                        || !string.Equals(pageWebMapping.TargetWebUrl.TrimEnd('/'), targetWeb.Url.TrimEnd('/'), StringComparison.OrdinalIgnoreCase))
                    {
                        blockers.Add("TargetPageWebTopologyMismatch: the target connection Web must be the mapped target for the source page Web. Supply a topology override or connect to the mapped child Web.");
                    }

                    result.ListMigration = ListMigrationPlanFactory.Create(
                        snapshot.ListDependencies,
                        snapshot.ListLookupDependencies,
                        result.Topology,
                        options.TaxonomySchemaMappings,
                        options.ListTargetOverrides);
                    var listTargetAnalysis = ListMigrationTargetAnalyzer.PopulateAndSeal(
                        targetContext,
                        snapshot.ListDependencies,
                        result.ListMigration,
                        result.TopologyTargetAnalysis);
                    AddIssues(listTargetAnalysis.Issues, blockers);
                    foreach (var warning in listTargetAnalysis.Warnings)
                    {
                        warnings.Add(warning);
                    }
                }
            }
            else if (snapshot.ListWebPartBindings.Count > 0 || snapshot.ListDependencies.Count > 0)
            {
                blockers.Add("SourceTopologyUnavailable: list-bound Web Parts require the exact source Web ownership closure.");
            }

            result.WebPartActions = ClassicWebPartActionPlanner.Build(
                snapshot.WebParts,
                snapshot.ListWebPartBindings,
                result.ListMigration,
                blockers);
            return result;
        }

        private static void AddIssues(
            IEnumerable<Diagnostics.MigrationIssue> issues,
            ICollection<string> blockers)
        {
            foreach (var issue in issues)
            {
                blockers.Add(issue.Code + ": " + issue.Message);
            }
        }
    }
}
