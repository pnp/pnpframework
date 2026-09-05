using Microsoft.SharePoint.Client;
using PnP.Framework.Migration.Diagnostics;
using PnP.Framework.Migration.Features;
using PnP.Framework.Migration.Lists.Capture;
using PnP.Framework.Migration.Schema.ContentTypes;
using PnP.Framework.Migration.Topology;
using System;
using System.Collections.Generic;
using System.Linq;

namespace PnP.Framework.Migration.Lists.Planning
{
    internal sealed class ListMigrationTargetAnalysisResult
    {
        public IList<MigrationIssue> Issues { get; set; } = new List<MigrationIssue>();

        public IList<string> Warnings { get; set; } = new List<string>();

        public bool IsAdmitted => Issues.All(value => value.Severity != MigrationIssueSeverity.Blocker
            && value.Severity != MigrationIssueSeverity.Error);
    }

    internal static class ListMigrationTargetAnalyzer
    {
        public static ListMigrationTargetAnalysisResult PopulateAndSeal(
            ClientContext targetContext,
            IEnumerable<ListDependencySnapshot> snapshots,
            ListMigrationPlanSet planSet,
            TopologyTargetAnalysis topologyAnalysis)
        {
            return Analyze(targetContext, snapshots, planSet, topologyAnalysis, true);
        }

        public static ListMigrationTargetAnalysisResult InspectFresh(
            ClientContext targetContext,
            IEnumerable<ListDependencySnapshot> snapshots,
            ListMigrationPlanSet planSet,
            TopologyTargetAnalysis topologyAnalysis)
        {
            return Analyze(targetContext, snapshots, planSet, topologyAnalysis, false);
        }

        private static ListMigrationTargetAnalysisResult Analyze(
            ClientContext targetContext,
            IEnumerable<ListDependencySnapshot> snapshots,
            ListMigrationPlanSet planSet,
            TopologyTargetAnalysis topologyAnalysis,
            bool populatePlan)
        {
            var result = new ListMigrationTargetAnalysisResult();
            var sources = (snapshots ?? Enumerable.Empty<ListDependencySnapshot>()).ToDictionary(value => value.SourceListId);
            if (sources.Count == 0 && planSet == null)
            {
                return result;
            }
            if (targetContext == null || planSet == null || topologyAnalysis == null)
            {
                result.Issues.Add(Issue("ListTargetAnalysisUnavailable", "target-lists",
                    "List target analysis requires a target connection, List plan set, and topology target analysis."));
                return result;
            }

            var topologyProbes = topologyAnalysis.SiteCollections.SelectMany(value => value.Webs)
                .ToDictionary(value => value.SourceWebId);
            foreach (var issue in planSet.Issues)
            {
                result.Issues.Add(issue);
            }
            foreach (var listPlan in planSet.Lists)
            {
                if (populatePlan)
                {
                    listPlan.TargetProbe = null;
                    foreach (var contentTypePlan in listPlan.SiteContentTypes)
                    {
                        contentTypePlan.TargetProbe = null;
                        contentTypePlan.TargetAdmission = null;
                        contentTypePlan.DeferredUntilTopologyMaterialization = false;
                    }
                    foreach (var featurePlan in listPlan.RequiredFeatures)
                    {
                        featurePlan.TargetProbe = null;
                    }
                }
                foreach (var issue in listPlan.Issues)
                {
                    result.Issues.Add(issue);
                }
                ListDependencySnapshot source;
                if (!sources.TryGetValue(listPlan.SourceListId, out source))
                {
                    result.Issues.Add(Issue("ListSourceEvidenceUnavailable", "list:" + listPlan.SourceListId.ToString("D"),
                        "The List plan has no matching source snapshot."));
                    continue;
                }
                if (listPlan.Issues.Any(value => value.Severity == MigrationIssueSeverity.Blocker || value.Severity == MigrationIssueSeverity.Error))
                {
                    continue;
                }

                TopologyWebTargetProbe ownerProbe;
                if (!topologyProbes.TryGetValue(listPlan.SourceWebId, out ownerProbe) || !ownerProbe.IsAdmitted)
                {
                    result.Issues.Add(Issue("TargetListOwnerWebBlocked", "list:" + listPlan.SourceListId.ToString("D"),
                        "The source List has no admitted target owner Web."));
                }
                else if (!ownerProbe.Exists && ownerProbe.Disposition == TopologyMaterializationDisposition.CreateOwned)
                {
                    if (populatePlan)
                    {
                        listPlan.TargetProbe = ListTargetInspector.DeferUntilTopologyMaterialization(listPlan);
                        foreach (var featurePlan in listPlan.RequiredFeatures)
                        {
                            featurePlan.TargetProbe = PlatformFeatureTargetInspector.DeferUntilTopologyMaterialization(featurePlan);
                        }
                    }
                }
                else
                {
                    var listProbe = populatePlan
                        ? ListTargetInspector.InspectForPlanning(targetContext, source, listPlan)
                        : ListTargetInspector.Inspect(targetContext, source, listPlan);
                    if (populatePlan)
                    {
                        if (!string.Equals(listPlan.TargetRootFolderServerRelativeUrl, listProbe.TargetRootFolderServerRelativeUrl, StringComparison.Ordinal)
                            || !string.Equals(listPlan.TargetTitle, listProbe.TargetTitle, StringComparison.Ordinal))
                        {
                            listPlan.TargetRootFolderServerRelativeUrl = listProbe.TargetRootFolderServerRelativeUrl;
                            listPlan.TargetTitle = listProbe.TargetTitle;
                            listPlan.PlanDigest = ListMigrationPlanFactory.ComputePlanDigest(listPlan);
                        }
                        listPlan.TargetProbe = listProbe;
                        if (listProbe.CollisionResolved)
                        {
                            result.Warnings.Add(listProbe.CollisionResolutionReason + " Final target: '"
                                + listProbe.TargetRootFolderServerRelativeUrl + "' (title '" + listProbe.TargetTitle + "').");
                        }
                    }
                    foreach (var issue in listProbe.Issues)
                    {
                        result.Issues.Add(issue);
                    }
                    if (ownerProbe.TargetWebId.HasValue
                        && listProbe.TargetWebId.HasValue
                        && ownerProbe.TargetWebId.Value != listProbe.TargetWebId.Value)
                    {
                        result.Issues.Add(Issue("TargetListOwnerIdentityMismatch", "list:" + listPlan.SourceListId.ToString("D"),
                            "The List target probe resolved a different target Web ID than the admitted topology mapping."));
                    }
                    AnalyzeFeatures(targetContext, listPlan, populatePlan, result);
                }

                var admissionContext = new ContentTypeTargetAdmissionContext(
                    ContentTypeRuntimeCatalog.RuntimeFieldIdsProvidedBy(
                        listPlan.RequiredFeatures.Where(value => value.TargetProbe == null || value.TargetProbe.IsAdmitted)));
                foreach (var contentTypePlan in listPlan.SiteContentTypes
                             .OrderBy(value => value.Schema.ContentTypeId.Length)
                             .ThenBy(value => value.Schema.ContentTypeId, StringComparer.OrdinalIgnoreCase))
                {
                    var admission = AnalyzeContentType(
                        targetContext,
                        topologyProbes,
                        contentTypePlan,
                        admissionContext,
                        populatePlan,
                        result);
                    if (admission?.IsEligible == true)
                    {
                        admissionContext.RegisterAdmitted(contentTypePlan.Schema);
                    }
                }
            }

            result.Issues = result.Issues
                .GroupBy(value => value.Code + "\u001f" + value.Subject + "\u001f" + value.Message, StringComparer.Ordinal)
                .Select(value => value.First())
                .OrderBy(value => value.Code, StringComparer.Ordinal)
                .ThenBy(value => value.Subject, StringComparer.Ordinal)
                .ToList();
            result.Warnings = result.Warnings.Distinct(StringComparer.Ordinal).OrderBy(value => value, StringComparer.Ordinal).ToList();
            if (populatePlan)
            {
                ListMigrationPlanFactory.SealTargetAnalysis(planSet);
            }
            return result;
        }

        private static void AnalyzeFeatures(
            ClientContext targetContext,
            ListMaterializationPlan listPlan,
            bool populatePlan,
            ListMigrationTargetAnalysisResult result)
        {
            if (listPlan.RequiredFeatures.Count == 0)
            {
                return;
            }
            var probes = PlatformFeatureTargetInspector.Inspect(targetContext, listPlan.RequiredFeatures);
            foreach (var featurePlan in listPlan.RequiredFeatures)
            {
                PlatformFeatureTargetProbe probe;
                if (!probes.TryGetValue(featurePlan.FeatureId, out probe))
                {
                    result.Issues.Add(Issue(
                        "TargetFeatureProbeUnavailable",
                        "platform-feature:" + featurePlan.FeatureId.ToString("D"),
                        "The target feature analyzer did not return a probe for the required feature."));
                    continue;
                }
                if (populatePlan)
                {
                    featurePlan.TargetProbe = probe;
                }
                foreach (var issue in probe.Issues)
                {
                    result.Issues.Add(issue);
                }
                if (probe.IsAdmitted && !probe.IsActive)
                {
                    result.Warnings.Add("Target site feature '" + featurePlan.Name + "' ("
                        + featurePlan.FeatureId.ToString("D") + ") will be activated before List materialization.");
                }
            }
        }

        private static ContentTypeTargetAdmission AnalyzeContentType(
            ClientContext targetContext,
            IDictionary<Guid, TopologyWebTargetProbe> topologyProbes,
            ContentTypeClosureNodePlan plan,
            ContentTypeTargetAdmissionContext admissionContext,
            bool populatePlan,
            ListMigrationTargetAnalysisResult result)
        {
            TopologyWebTargetProbe ownerProbe;
            if (!topologyProbes.TryGetValue(plan.SourceOwnerWebId, out ownerProbe) || !ownerProbe.IsAdmitted)
            {
                result.Issues.Add(Issue("TargetContentTypeOwnerWebBlocked", "content-type:" + plan.Schema.ContentTypeId,
                    "The site content type has no admitted target owner Web."));
                if (populatePlan)
                {
                    plan.TargetAdmission = new ContentTypeTargetAdmission
                    {
                        Disposition = ContentTypeMaterializationDisposition.Block,
                        IsEligible = false
                    };
                }
                return plan.TargetAdmission;
            }
            if (!ownerProbe.Exists && ownerProbe.Disposition == TopologyMaterializationDisposition.CreateOwned)
            {
                if (populatePlan)
                {
                    plan.DeferredUntilTopologyMaterialization = true;
                }
                return null;
            }

            ContentTypeTargetProbe probe;
            ContentTypeTargetAdmission admission;
            using (var contentTypeContext = targetContext.Clone(plan.TargetOwnerWebUrl))
            {
                probe = ContentTypeTargetInspector.Inspect(contentTypeContext, contentTypeContext.Web, plan.Schema);
                admission = ContentTypeTargetAdmissionEvaluator.Evaluate(plan.Schema, probe, admissionContext);
            }
            if (populatePlan)
            {
                plan.TargetProbe = probe;
                plan.TargetAdmission = admission;
                plan.DeferredUntilTopologyMaterialization = false;
            }
            foreach (var issue in admission.Issues)
            {
                result.Issues.Add(issue);
            }
            foreach (var warning in admission.Warnings)
            {
                result.Warnings.Add(warning);
            }
            return admission;
        }

        private static MigrationIssue Issue(string code, string subject, string message)
        {
            return new MigrationIssue
            {
                Code = code,
                Severity = MigrationIssueSeverity.Blocker,
                Subject = subject,
                Ingredient = "ListDependency.Target",
                Message = message
            };
        }
    }
}
