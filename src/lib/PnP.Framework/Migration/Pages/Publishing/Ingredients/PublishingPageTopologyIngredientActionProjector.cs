using PnP.Framework.Migration.Diagnostics;
using PnP.Framework.Migration.Pages.Ingredients;
using PnP.Framework.Migration.Pages.Publishing.Capture;
using PnP.Framework.Migration.Pages.Publishing.Planning;
using PnP.Framework.Migration.Topology;
using System;
using System.Collections.Generic;
using System.Linq;

namespace PnP.Framework.Migration.Pages.Publishing.Ingredients
{
    internal static class PublishingPageTopologyIngredientActionProjector
    {
        public static void Project(
            PublishingPageCaptureBundle snapshot,
            PublishingPageMigrationPlan plan,
            IDictionary<string, PageIngredientAction> actions)
        {
            if (snapshot.SourceTopology == null)
            {
                return;
            }

            var mappings = (plan.Topology?.SiteCollections ?? Array.Empty<SiteCollectionMappingPlan>())
                .SelectMany(value => value.Webs ?? Array.Empty<WebMappingPlan>())
                .GroupBy(value => WebKey(value.SourceSiteId, value.SourceWebId), StringComparer.Ordinal)
                .ToDictionary(group => group.Key, group => group.First(), StringComparer.Ordinal);
            var probes = (plan.TopologyTargetAnalysis?.SiteCollections ?? Array.Empty<TopologySiteTargetProbe>())
                .SelectMany(value => value.Webs ?? Array.Empty<TopologyWebTargetProbe>())
                .GroupBy(value => WebKey(value.SourceSiteId, value.SourceWebId), StringComparer.Ordinal)
                .ToDictionary(group => group.Key, group => group.First(), StringComparer.Ordinal);

            foreach (var sourceWeb in snapshot.SourceTopology.Webs.Where(value => value != null))
            {
                var key = WebKey(sourceWeb.SiteId, sourceWeb.WebId);
                mappings.TryGetValue(key, out var mapping);
                probes.TryGetValue(key, out var probe);
                var blocked = mapping == null
                    || probe == null
                    || probe.Disposition == TopologyMaterializationDisposition.Block
                    || !probe.IsAdmitted;
                var reason = mapping == null
                    ? "The captured source Web has no target topology mapping."
                    : probe == null
                        ? "The mapped target Web has no sealed fresh target analysis."
                        : blocked
                            ? "The mapped target Web failed topology admission: " + JoinIssues(probe.Issues)
                            : "Materialize or reuse the source Web at its reviewed target topology identity.";
                PublishingPageIngredientActionFactory.Add(actions, PublishingPageIngredientActionFactory.Create(
                    PublishingPageIngredientIds.Web(sourceWeb.SiteId, sourceWeb.WebId),
                    blocked ? IngredientCapability.Incompatible : IngredientCapability.Available,
                    blocked ? IngredientDisposition.Block : IngredientDisposition.Preserve,
                    blocked ? "none" : Realization(probe.Disposition),
                    "policy.topology.web",
                    reason,
                    mapping?.TargetWebUrl,
                    blocked
                        ? null
                        : $"The topology receipt maps source Web '{sourceWeb.WebId:D}' to '{mapping.TargetWebUrl}'.",
                    blocked ? null : "Fresh readback confirms the target Web identity, parent, template, and ownership marker."));
            }
        }

        private static string Realization(TopologyMaterializationDisposition disposition)
        {
            switch (disposition)
            {
                case TopologyMaterializationDisposition.CreateOwned:
                    return "create-owned";
                case TopologyMaterializationDisposition.ReuseOwned:
                    return "reuse-owned";
                case TopologyMaterializationDisposition.ReuseApprovedHost:
                    return "reuse-approved-host";
                case TopologyMaterializationDisposition.RecoverInterruptedCreate:
                    return "recover-interrupted-create";
                default:
                    return "none";
            }
        }

        private static string WebKey(Guid sourceSiteId, Guid sourceWebId)
        {
            return sourceSiteId.ToString("D") + "/" + sourceWebId.ToString("D");
        }

        private static string JoinIssues(IEnumerable<MigrationIssue> issues)
        {
            var messages = (issues ?? Array.Empty<MigrationIssue>())
                .Select(value => value.Code + ": " + value.Message)
                .ToArray();
            return messages.Length == 0 ? "admission was denied" : string.Join("; ", messages);
        }
    }
}
