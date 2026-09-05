using PnP.Framework.Migration.Diagnostics;
using System;
using System.Linq;

namespace PnP.Framework.Migration.Topology
{
    internal static class TopologyTargetInspector
    {
        public static TopologyTargetAnalysis Inspect(
            Microsoft.SharePoint.Client.ClientContext anchorContext,
            TopologyPlan plan,
            string approvedHostWebUrl)
        {
            return Inspect(anchorContext, plan, approvedHostWebUrl, false);
        }

        public static TopologyTargetAnalysis InspectForPlanning(
            Microsoft.SharePoint.Client.ClientContext anchorContext,
            TopologyPlan plan,
            string approvedHostWebUrl)
        {
            return Inspect(anchorContext, plan, approvedHostWebUrl, true);
        }

        private static TopologyTargetAnalysis Inspect(
            Microsoft.SharePoint.Client.ClientContext anchorContext,
            TopologyPlan plan,
            string approvedHostWebUrl,
            bool resolvePlanningCollisions)
        {
            if (anchorContext == null)
            {
                throw new ArgumentNullException(nameof(anchorContext));
            }
            if (plan == null)
            {
                throw new ArgumentNullException(nameof(plan));
            }

            var scope = TopologyTargetInspectionScope.Create(anchorContext, approvedHostWebUrl);
            var result = new TopologyTargetAnalysis { TopologyPlanDigest = plan.PlanDigest };
            foreach (var sitePlan in plan.SiteCollections)
            {
                var siteProbe = resolvePlanningCollisions
                    ? TopologySiteTargetInspector.InspectForPlanning(scope, sitePlan)
                    : TopologySiteTargetInspector.Inspect(scope, sitePlan);
                result.SiteCollections.Add(siteProbe);
                foreach (var issue in siteProbe.Issues.Concat(siteProbe.Webs.SelectMany(value => value.Issues)))
                {
                    result.Issues.Add(issue);
                }
            }

            result.SiteCollections = result.SiteCollections.OrderBy(value => value.SourceSiteId).ToList();
            result.Issues = result.Issues.OrderBy(value => value.Code, StringComparer.Ordinal)
                .ThenBy(value => value.Subject, StringComparer.Ordinal).ToList();
            if (resolvePlanningCollisions)
            {
                plan.PlanDigest = TopologyPlanner.ComputeDigest(plan);
                result.TopologyPlanDigest = plan.PlanDigest;
            }
            return result;
        }

        internal static string InterruptedCreateDescription(WebMappingPlan plan)
        {
            return TopologyWebTargetInspector.InterruptedCreateDescription(plan);
        }
    }
}
