using Microsoft.SharePoint.Client;
using PnP.Framework.Migration.Diagnostics;
using System;
using System.Collections.Generic;
using System.Linq;
using static PnP.Framework.Migration.Topology.TopologyTargetInspectionScope;

namespace PnP.Framework.Migration.Topology
{
    internal static class TopologySiteTargetInspector
    {
        public static TopologySiteTargetProbe Inspect(TopologyTargetInspectionScope scope, SiteCollectionMappingPlan plan)
        {
            return Inspect(scope, plan, false);
        }

        public static TopologySiteTargetProbe InspectForPlanning(TopologyTargetInspectionScope scope, SiteCollectionMappingPlan plan)
        {
            return Inspect(scope, plan, true);
        }

        private static TopologySiteTargetProbe Inspect(
            TopologyTargetInspectionScope scope,
            SiteCollectionMappingPlan plan,
            bool resolvePlanningCollisions)
        {
            var probe = new TopologySiteTargetProbe
            {
                SourceSiteId = plan.SourceSiteId,
                PreferredTargetSiteCollectionUrl = plan.PreferredTargetSiteCollectionUrl ?? plan.TargetSiteCollectionUrl,
                TargetSiteCollectionUrl = plan.TargetSiteCollectionUrl,
                CollisionResolved = plan.TargetSiteCollisionResolved,
                CollisionResolutionReason = plan.TargetSiteResolutionReason,
                Disposition = TopologyMaterializationDisposition.Block
            };
            if (plan.TargetMode != TargetSiteMode.ExistingTargetSite)
            {
                probe.Issues.Add(Issue("TargetSiteCreationRequiresTenantExecutor", "target-site:" + plan.TargetSiteCollectionUrl,
                    "This importer has a Web-scoped target connection and cannot create a new Site Collection."));
                return probe;
            }

            try
            {
                if (scope.LoadedRoot != null && TargetUrl.Equals(plan.TargetSiteCollectionUrl, scope.LoadedRoot.Url))
                {
                    return Populate(scope, plan, probe, scope.LoadedRoot, resolvePlanningCollisions);
                }

                using (var context = scope.AnchorContext.Clone(plan.TargetSiteCollectionUrl))
                {
                    var site = context.Site;
                    var root = site.RootWeb;
                    context.Load(site, value => value.Id);
                    context.Load(root,
                        value => value.Id,
                        value => value.Url,
                        value => value.Title,
                        value => value.WebTemplate,
                        value => value.Configuration);
                    context.ExecuteQueryRetry();
                    return Populate(scope, plan, probe, new LoadedRootTarget(
                        site.Id,
                        root.Id,
                        root.Url,
                        root.Title,
                        root.WebTemplate,
                        root.Configuration), resolvePlanningCollisions);
                }
            }
            catch (Exception exception) when (exception is ServerException || exception is ClientRequestException)
            {
                probe.Issues.Add(Issue("TargetSiteUnavailable", "target-site:" + plan.TargetSiteCollectionUrl,
                    "The planned target Site Collection could not be inspected: " + exception.Message));
                return probe;
            }
        }

        private static TopologySiteTargetProbe Populate(
            TopologyTargetInspectionScope scope,
            SiteCollectionMappingPlan plan,
            TopologySiteTargetProbe probe,
            LoadedRootTarget root,
            bool resolvePlanningCollisions)
        {
            probe.Exists = true;
            probe.TargetSiteId = root.SiteId;
            probe.TargetRootWebId = root.WebId;
            if (plan.ExpectedTargetSiteId.HasValue && root.SiteId != plan.ExpectedTargetSiteId.Value)
            {
                probe.Issues.Add(Issue("TargetSiteIdentityMismatch", "target-site:" + plan.TargetSiteCollectionUrl,
                    "Expected target Site Collection " + plan.ExpectedTargetSiteId.Value.ToString("D") + ", observed " + root.SiteId.ToString("D") + "."));
                return probe;
            }
            if (scope.ApprovedSiteId != root.SiteId)
            {
                probe.Issues.Add(Issue("ApprovedHostSiteMismatch", "target-site:" + plan.TargetSiteCollectionUrl,
                    "The explicit target connection is outside the planned target Site Collection."));
                return probe;
            }

            var rootPlan = plan.Webs.SingleOrDefault(value => value.Kind == TopologyNodeKind.SiteCollectionRoot);
            if (rootPlan == null || !TargetUrl.Equals(rootPlan.TargetWebUrl, root.Url))
            {
                probe.Issues.Add(Issue("TargetRootMappingMismatch", "target-site:" + plan.TargetSiteCollectionUrl,
                    "The topology plan has no exact mapping for the observed target root Web."));
                return probe;
            }

            probe.Disposition = TopologyMaterializationDisposition.ReuseApprovedHost;
            var probes = new Dictionary<Guid, TopologyWebTargetProbe>
            {
                [rootPlan.SourceWebId] = new TopologyWebTargetProbe
                {
                    SourceSiteId = rootPlan.SourceSiteId,
                    SourceWebId = rootPlan.SourceWebId,
                    PreferredTargetWebUrl = rootPlan.PreferredTargetWebUrl ?? rootPlan.TargetWebUrl,
                    TargetWebUrl = rootPlan.TargetWebUrl,
                    PreferredTargetServerRelativeUrl = rootPlan.PreferredTargetServerRelativeUrl ?? rootPlan.TargetServerRelativeUrl,
                    TargetServerRelativeUrl = rootPlan.TargetServerRelativeUrl,
                    Exists = true,
                    TargetSiteId = root.SiteId,
                    TargetWebId = root.WebId,
                    ExistingTitle = root.Title,
                    ExistingTemplate = root.Template,
                    ExistingConfiguration = root.Configuration,
                    Disposition = TopologyMaterializationDisposition.ReuseApprovedHost
                }
            };

            foreach (var childPlan in plan.Webs.Where(value => value.Kind == TopologyNodeKind.ChildWeb)
                         .OrderBy(value => TargetUrl.Depth(value.TargetServerRelativeUrl))
                         .ThenBy(value => value.TargetServerRelativeUrl, StringComparer.OrdinalIgnoreCase)
                         .ToArray())
            {
                TopologyWebTargetProbe parentProbe;
                if (!childPlan.SourceParentWebId.HasValue || !probes.TryGetValue(childPlan.SourceParentWebId.Value, out parentProbe))
                {
                    probes[childPlan.SourceWebId] = TopologyWebTargetInspector.Blocked(
                        childPlan,
                        "TargetParentMappingUnavailable",
                        "The target parent Web mapping is unavailable.");
                    continue;
                }
                if (!parentProbe.IsAdmitted)
                {
                    probes[childPlan.SourceWebId] = TopologyWebTargetInspector.Blocked(
                        childPlan,
                        "TargetParentBlocked",
                        "The target parent Web is blocked.");
                    continue;
                }
                if (!parentProbe.Exists)
                {
                    probes[childPlan.SourceWebId] = TopologyWebTargetInspector.PlannedCreate(
                        childPlan,
                        root.SiteId,
                        parentProbe.TargetWebId);
                    continue;
                }

                var childProbe = resolvePlanningCollisions
                    ? TopologyWebTargetInspector.InspectForPlanning(
                        scope,
                        childPlan,
                        root.SiteId,
                        parentProbe.TargetWebId.Value)
                    : TopologyWebTargetInspector.Inspect(
                        scope,
                        childPlan,
                        root.SiteId,
                        parentProbe.TargetWebId.Value);
                if (resolvePlanningCollisions && childProbe.CollisionResolved)
                {
                    TopologyPlanRetargeter.RetargetWeb(
                        plan,
                        childPlan.SourceWebId,
                        childProbe.TargetServerRelativeUrl);
                }
                probes[childPlan.SourceWebId] = childProbe;
            }

            probe.Webs = plan.Webs.Select(value => probes[value.SourceWebId]).ToList();
            return probe;
        }

        private static MigrationIssue Issue(string code, string subject, string message)
        {
            return new MigrationIssue
            {
                Code = code,
                Severity = MigrationIssueSeverity.Blocker,
                Subject = subject,
                Ingredient = "Topology.Target",
                Message = message
            };
        }
    }
}
