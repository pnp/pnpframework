using Microsoft.SharePoint.Client;
using PnP.Framework.Migration.Execution;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;

namespace PnP.Framework.Migration.Topology.Execution
{
    internal static class TopologyMaterializationCoordinator
    {
        public static TopologyMaterializationReceipt Ensure(
            ClientContext anchorContext,
            TopologyPlan plan,
            MigrationExecutionRecorder recorder)
        {
            return Ensure(anchorContext, plan, plan, recorder);
        }

        public static TopologyMaterializationReceipt Ensure(
            ClientContext anchorContext,
            TopologyPlan approvedPlan,
            TopologyPlan executionPlan,
            MigrationExecutionRecorder recorder)
        {
            if (executionPlan == null)
            {
                recorder.RecordAlreadySatisfied("topology.materialize", "The approved package has no source topology to materialize.");
                return new TopologyMaterializationReceipt
                {
                    FreshReadbackPassed = true,
                    TopologyPlanDigest = approvedPlan?.PlanDigest,
                    ApprovedTopologyPlanDigest = approvedPlan?.PlanDigest,
                    Diagnostics = new List<string> { "No topology closure was required." }
                };
            }
            var plan = executionPlan;

            if (!anchorContext.Web.IsPropertyAvailable("Url"))
            {
                anchorContext.Load(anchorContext.Web, value => value.Url);
                anchorContext.ExecuteQueryRetry();
            }
            var approvedHostUrl = anchorContext.Web.Url;
            var initial = TopologyTargetInspector.Inspect(anchorContext, plan, approvedHostUrl);
            if (!initial.IsAdmitted)
            {
                throw new InvalidOperationException("Fresh target topology preflight failed: "
                    + string.Join("; ", initial.Issues.Select(value => value.Message)));
            }

            var result = new TopologyMaterializationReceipt
            {
                TopologyPlanDigest = approvedPlan?.PlanDigest ?? plan.PlanDigest,
                ApprovedTopologyPlanDigest = approvedPlan?.PlanDigest,
                ExecutionTopologyPlanDigest = plan.PlanDigest
            };
            if (TryCompleteWithoutMutation(plan, initial, recorder, result))
            {
                return result;
            }

            var inspectionScope = TopologyTargetInspectionScope.Create(anchorContext, approvedHostUrl);
            var initialSites = initial.SiteCollections.ToDictionary(value => value.SourceSiteId);
            foreach (var sitePlan in plan.SiteCollections.OrderBy(value => value.SourceSiteId))
            {
                TopologySiteTargetProbe initialSite;
                if (!initialSites.TryGetValue(sitePlan.SourceSiteId, out initialSite)
                    || !initialSite.TargetSiteId.HasValue)
                {
                    throw new InvalidDataException(
                        "The admitted target analysis has no runtime Site identity for source Site "
                        + sitePlan.SourceSiteId.ToString("D") + ".");
                }
                var initialWebs = initialSite.Webs.ToDictionary(value => value.SourceWebId);
                var materializedProbes = new Dictionary<Guid, TopologyWebTargetProbe>();
                foreach (var webPlan in sitePlan.Webs.OrderBy(value => Depth(value.TargetServerRelativeUrl))
                             .ThenBy(value => value.TargetServerRelativeUrl, StringComparer.OrdinalIgnoreCase))
                {
                    TopologyWebTargetProbe probe;
                    if (webPlan.Kind == TopologyNodeKind.SiteCollectionRoot)
                    {
                        if (!initialWebs.TryGetValue(webPlan.SourceWebId, out probe))
                        {
                            throw new InvalidDataException(
                                "The admitted target analysis has no root-Web probe for source Web "
                                + webPlan.SourceWebId.ToString("D") + ".");
                        }
                    }
                    else
                    {
                        TopologyWebTargetProbe parentProbe;
                        if (!webPlan.SourceParentWebId.HasValue
                            || !materializedProbes.TryGetValue(webPlan.SourceParentWebId.Value, out parentProbe)
                            || !parentProbe.TargetSiteId.HasValue
                            || !parentProbe.TargetWebId.HasValue)
                        {
                            throw new InvalidDataException(
                                "The direct target parent has not been materialized for source Web "
                                + webPlan.SourceWebId.ToString("D") + ".");
                        }
                        probe = TopologyWebTargetInspector.Inspect(
                            inspectionScope,
                            webPlan,
                            parentProbe.TargetSiteId.Value,
                            parentProbe.TargetWebId.Value);
                    }
                    if (!probe.IsAdmitted)
                    {
                        throw new InvalidOperationException("Fresh target Web preflight failed for '" + webPlan.TargetWebUrl + "': "
                            + string.Join("; ", probe.Issues.Select(value => value.Message)));
                    }

                    var disposition = probe.Disposition;
                    if (webPlan.Kind == TopologyNodeKind.SiteCollectionRoot
                        || disposition == TopologyMaterializationDisposition.ReuseApprovedHost
                        || disposition == TopologyMaterializationDisposition.ReuseOwned)
                    {
                        recorder.RecordAlreadySatisfied("topology.web." + webPlan.SourceWebId.ToString("N"),
                            "Reuse target Web '" + webPlan.TargetWebUrl + "' as " + disposition + ".");
                    }
                    else
                    {
                        probe = recorder.Execute(
                            "topology.web." + webPlan.SourceWebId.ToString("N"),
                            disposition == TopologyMaterializationDisposition.RecoverInterruptedCreate
                                ? "Recover and claim interrupted target Web '" + webPlan.TargetWebUrl + "'."
                                : "Create and claim target Web '" + webPlan.TargetWebUrl + "'.",
                            () => EnsureChild(
                                anchorContext,
                                inspectionScope,
                                sitePlan,
                                webPlan,
                                probe),
                            value => MutationOutcome.Applied,
                            value => "Target Web " + value.TargetWebId.Value.ToString("D") + " passed exact provenance readback.");
                    }

                    if (!probe.TargetWebId.HasValue || !probe.TargetSiteId.HasValue)
                    {
                        throw new InvalidDataException("Materialized topology Web has no runtime Site/Web identity: " + webPlan.TargetWebUrl + ".");
                    }
                    result.Webs.Add(new TopologyWebMaterializationReceipt
                    {
                        SourceSiteId = webPlan.SourceSiteId,
                        SourceWebId = webPlan.SourceWebId,
                        TargetSiteId = probe.TargetSiteId.Value,
                        TargetWebId = probe.TargetWebId.Value,
                        TargetWebUrl = webPlan.TargetWebUrl,
                        Disposition = disposition,
                        MappingDigest = TopologyPlanner.ComputeWebMappingDigest(webPlan)
                    });
                    materializedProbes[webPlan.SourceWebId] = probe;
                }
            }
            var final = TopologyTargetInspector.Inspect(anchorContext, plan, approvedHostUrl);
            var finalWebs = final.SiteCollections.SelectMany(value => value.Webs).ToDictionary(value => value.SourceWebId);
            var mismatches = new List<string>();
            if (!final.IsAdmitted || finalWebs.Count != result.Webs.Count)
            {
                mismatches.Add("Final topology target analysis is blocked or does not cover every materialized Web.");
            }
            foreach (var receipt in result.Webs)
            {
                TopologyWebTargetProbe probe;
                if (!finalWebs.TryGetValue(receipt.SourceWebId, out probe)
                    || !probe.TargetSiteId.HasValue
                    || !probe.TargetWebId.HasValue
                    || probe.TargetSiteId.Value != receipt.TargetSiteId
                    || probe.TargetWebId.Value != receipt.TargetWebId
                    || !UrlEquals(probe.TargetWebUrl, receipt.TargetWebUrl)
                    || (probe.Disposition != TopologyMaterializationDisposition.ReuseApprovedHost
                        && probe.Disposition != TopologyMaterializationDisposition.ReuseOwned))
                {
                    mismatches.Add("Final topology readback differs for source Web " + receipt.SourceWebId.ToString("D") + ".");
                }
            }
            if (mismatches.Count > 0)
            {
                throw new InvalidOperationException(string.Join("; ", mismatches));
            }
            result.FreshReadbackPassed = true;
            result.Diagnostics.Add("Fresh target analysis verified " + result.Webs.Count + " Site/Web mapping(s) after topology materialization.");
            return result;
        }

        internal static bool TryCompleteWithoutMutation(
            TopologyPlan plan,
            TopologyTargetAnalysis analysis,
            MigrationExecutionRecorder recorder,
            TopologyMaterializationReceipt result)
        {
            var plans = plan.SiteCollections
                .SelectMany(value => value.Webs)
                .OrderBy(value => Depth(value.TargetServerRelativeUrl))
                .ThenBy(value => value.TargetServerRelativeUrl, StringComparer.OrdinalIgnoreCase)
                .ToArray();
            var probes = analysis.SiteCollections
                .SelectMany(value => value.Webs)
                .ToDictionary(value => value.SourceWebId);
            if (plans.Length != probes.Count)
            {
                return false;
            }

            foreach (var webPlan in plans)
            {
                if (!probes.TryGetValue(webPlan.SourceWebId, out var probe)
                    || !probe.IsAdmitted
                    || !probe.TargetWebId.HasValue
                    || !probe.TargetSiteId.HasValue
                    || (probe.Disposition != TopologyMaterializationDisposition.ReuseApprovedHost
                        && probe.Disposition != TopologyMaterializationDisposition.ReuseOwned))
                {
                    return false;
                }
            }

            foreach (var webPlan in plans)
            {
                var probe = probes[webPlan.SourceWebId];
                recorder.RecordAlreadySatisfied(
                    "topology.web." + webPlan.SourceWebId.ToString("N"),
                    "Reuse target Web '" + webPlan.TargetWebUrl + "' as " + probe.Disposition + ".");
                result.Webs.Add(new TopologyWebMaterializationReceipt
                {
                    SourceSiteId = webPlan.SourceSiteId,
                    SourceWebId = webPlan.SourceWebId,
                    TargetSiteId = probe.TargetSiteId.Value,
                    TargetWebId = probe.TargetWebId.Value,
                    TargetWebUrl = webPlan.TargetWebUrl,
                    Disposition = probe.Disposition,
                    MappingDigest = TopologyPlanner.ComputeWebMappingDigest(webPlan)
                });
            }

            result.FreshReadbackPassed = true;
            result.Diagnostics.Add(
                "One fresh target analysis verified " + result.Webs.Count
                + " reusable Site/Web mapping(s); no topology mutation was required.");
            return true;
        }

        private static TopologyWebTargetProbe EnsureChild(
            ClientContext anchorContext,
            TopologyTargetInspectionScope inspectionScope,
            SiteCollectionMappingPlan sitePlan,
            WebMappingPlan plan,
            TopologyWebTargetProbe admittedProbe)
        {
            if (plan.Kind != TopologyNodeKind.ChildWeb || string.IsNullOrWhiteSpace(plan.TargetParentWebUrl))
            {
                throw new InvalidOperationException("Only a planned child Web can be created or recovered.");
            }
            if (!admittedProbe.TargetSiteId.HasValue || !admittedProbe.TargetParentWebId.HasValue)
            {
                throw new InvalidDataException("The admitted target child-Web probe has no runtime Site/parent identity.");
            }

            using (var parentContext = anchorContext.Clone(plan.TargetParentWebUrl))
            {
                var parent = parentContext.Web;
                parentContext.Load(parent, value => value.Id, value => value.Url, value => value.ServerRelativeUrl);
                parentContext.ExecuteQueryRetry();
                if (!UrlEquals(parent.Url, plan.TargetParentWebUrl))
                {
                    throw new InvalidOperationException("The target connection did not resolve the planned direct parent Web.");
                }

                Web target;
                if (admittedProbe.Disposition == TopologyMaterializationDisposition.CreateOwned)
                {
                    var parentPath = parent.ServerRelativeUrl.TrimEnd('/');
                    if (!plan.TargetServerRelativeUrl.StartsWith(parentPath + "/", StringComparison.OrdinalIgnoreCase))
                    {
                        throw new InvalidDataException("The planned child Web path is outside its target parent Web.");
                    }
                    var relativePath = Uri.UnescapeDataString(plan.TargetServerRelativeUrl.Substring(parentPath.Length + 1));
                    if (string.IsNullOrWhiteSpace(relativePath)
                        || relativePath.IndexOf('/') >= 0
                        || relativePath.IndexOf('\\') >= 0
                        || relativePath == "."
                        || relativePath == "..")
                    {
                        throw new InvalidDataException("The planned child Web URL must be one safe segment below its captured direct parent.");
                    }
                    target = parent.Webs.Add(new WebCreationInformation
                    {
                        Url = relativePath,
                        Title = plan.TargetTitle,
                        Description = TopologyTargetInspector.InterruptedCreateDescription(plan),
                        Language = sitePlan.TargetLanguage <= 0 ? 1033 : sitePlan.TargetLanguage,
                        UseSamePermissionsAsParentSite = true,
                        WebTemplate = NormalizeTemplate(plan.TargetTemplate, plan.TargetConfiguration)
                    });
                    parentContext.Load(target, value => value.Id, value => value.Url, value => value.AllProperties);
                    parentContext.ExecuteQueryRetry();
                }
                else if (admittedProbe.Disposition == TopologyMaterializationDisposition.RecoverInterruptedCreate)
                {
                    if (!admittedProbe.TargetWebId.HasValue)
                    {
                        throw new InvalidDataException("The interrupted target child-Web probe has no runtime Web identity.");
                    }
                    target = parentContext.Site.OpenWebById(admittedProbe.TargetWebId.Value);
                    parentContext.Load(target, value => value.Id, value => value.Url, value => value.AllProperties);
                    parentContext.ExecuteQueryRetry();
                }
                else
                {
                    throw new InvalidOperationException("Unexpected child Web materialization disposition: " + admittedProbe.Disposition + ".");
                }

                target.AllProperties[TopologyPlanner.WebOriginalIdentifierPropertyName] = plan.OriginalIdentifier;
                target.AllProperties[TopologyPlanner.WebPlanDigestPropertyName] = TopologyPlanner.ComputeWebMappingDigest(plan);
                target.Update();
                parentContext.ExecuteQueryRetry();
            }

            var readback = TopologyWebTargetInspector.Inspect(
                inspectionScope,
                plan,
                admittedProbe.TargetSiteId.Value,
                admittedProbe.TargetParentWebId.Value);
            if (readback.Disposition != TopologyMaterializationDisposition.ReuseOwned)
            {
                throw new InvalidOperationException("Fresh target child-Web readback did not resolve exact migration ownership.");
            }
            return readback;
        }

        private static string NormalizeTemplate(string template, int configuration)
        {
            return (template ?? string.Empty).IndexOf('#') >= 0
                ? template
                : template + "#" + configuration.ToString(CultureInfo.InvariantCulture);
        }

        private static bool UrlEquals(string left, string right)
        {
            return string.Equals(new Uri(left).AbsoluteUri.TrimEnd('/'), new Uri(right).AbsoluteUri.TrimEnd('/'), StringComparison.OrdinalIgnoreCase);
        }

        private static int Depth(string value)
        {
            return (value ?? string.Empty).Count(character => character == '/');
        }
    }
}
