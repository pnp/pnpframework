using Microsoft.SharePoint.Client;
using PnP.Framework.Migration.Diagnostics;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;

namespace PnP.Framework.Migration.Topology
{
    internal static class TopologyTargetInspector
    {
        public static TopologyTargetAnalysis Inspect(ClientContext anchorContext, TopologyPlan plan, string approvedHostWebUrl)
        {
            if (anchorContext == null)
            {
                throw new ArgumentNullException(nameof(anchorContext));
            }
            if (plan == null)
            {
                throw new ArgumentNullException(nameof(plan));
            }

            anchorContext.Load(anchorContext.Site, value => value.Id);
            anchorContext.Load(anchorContext.Web, value => value.Id, value => value.Url);
            anchorContext.ExecuteQueryRetry();
            var approvedHostId = anchorContext.Web.Id;
            var approvedSiteId = anchorContext.Site.Id;
            var approvedUrl = NormalizeAbsolute(approvedHostWebUrl ?? anchorContext.Web.Url);
            var result = new TopologyTargetAnalysis { TopologyPlanDigest = plan.PlanDigest };
            foreach (var sitePlan in plan.SiteCollections)
            {
                var siteProbe = InspectSite(anchorContext, sitePlan, approvedSiteId, approvedHostId, approvedUrl);
                result.SiteCollections.Add(siteProbe);
                foreach (var issue in siteProbe.Issues.Concat(siteProbe.Webs.SelectMany(value => value.Issues)))
                {
                    result.Issues.Add(issue);
                }
            }
            result.SiteCollections = result.SiteCollections.OrderBy(value => value.SourceSiteId).ToList();
            result.Issues = result.Issues.OrderBy(value => value.Code, StringComparer.Ordinal)
                .ThenBy(value => value.Subject, StringComparer.Ordinal).ToList();
            return result;
        }

        private static TopologySiteTargetProbe InspectSite(
            ClientContext anchorContext,
            SiteCollectionMappingPlan plan,
            Guid approvedSiteId,
            Guid approvedHostId,
            string approvedHostUrl)
        {
            var probe = new TopologySiteTargetProbe
            {
                SourceSiteId = plan.SourceSiteId,
                TargetSiteCollectionUrl = plan.TargetSiteCollectionUrl,
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
                using (var context = anchorContext.Clone(plan.TargetSiteCollectionUrl))
                {
                    var site = context.Site;
                    var root = site.RootWeb;
                    context.Load(site, value => value.Id);
                    context.Load(root,
                        value => value.Id,
                        value => value.Url,
                        value => value.ServerRelativeUrl,
                        value => value.Title,
                        value => value.WebTemplate,
                        value => value.Configuration);
                    context.ExecuteQueryRetry();
                    probe.Exists = true;
                    probe.TargetSiteId = site.Id;
                    probe.TargetRootWebId = root.Id;
                    if (plan.ExpectedTargetSiteId.HasValue && site.Id != plan.ExpectedTargetSiteId.Value)
                    {
                        probe.Issues.Add(Issue("TargetSiteIdentityMismatch", "target-site:" + plan.TargetSiteCollectionUrl,
                            "Expected target Site Collection " + plan.ExpectedTargetSiteId.Value.ToString("D") + ", observed " + site.Id.ToString("D") + "."));
                        return probe;
                    }
                    if (approvedSiteId != site.Id)
                    {
                        probe.Issues.Add(Issue("ApprovedHostSiteMismatch", "target-site:" + plan.TargetSiteCollectionUrl,
                            "The explicit target connection is outside the planned target Site Collection."));
                        return probe;
                    }

                    var rootPlan = plan.Webs.SingleOrDefault(value => value.Kind == TopologyNodeKind.SiteCollectionRoot);
                    if (rootPlan == null || !UrlEquals(rootPlan.TargetWebUrl, root.Url))
                    {
                        probe.Issues.Add(Issue("TargetRootMappingMismatch", "target-site:" + plan.TargetSiteCollectionUrl,
                            "The topology plan has no exact mapping for the observed target root Web."));
                        return probe;
                    }
                    probe.Disposition = TopologyMaterializationDisposition.ReuseApprovedHost;
                    var probes = new Dictionary<Guid, TopologyWebTargetProbe>();
                    probes[rootPlan.SourceWebId] = new TopologyWebTargetProbe
                    {
                        SourceSiteId = rootPlan.SourceSiteId,
                        SourceWebId = rootPlan.SourceWebId,
                        TargetWebUrl = rootPlan.TargetWebUrl,
                        Exists = true,
                        TargetSiteId = site.Id,
                        TargetWebId = root.Id,
                        ExistingTitle = root.Title,
                        ExistingTemplate = root.WebTemplate,
                        ExistingConfiguration = root.Configuration,
                        Disposition = TopologyMaterializationDisposition.ReuseApprovedHost
                    };

                    foreach (var childPlan in plan.Webs.Where(value => value.Kind == TopologyNodeKind.ChildWeb)
                                 .OrderBy(value => Depth(value.TargetServerRelativeUrl))
                                 .ThenBy(value => value.TargetServerRelativeUrl, StringComparer.OrdinalIgnoreCase))
                    {
                        TopologyWebTargetProbe parentProbe;
                        if (!childPlan.SourceParentWebId.HasValue || !probes.TryGetValue(childPlan.SourceParentWebId.Value, out parentProbe))
                        {
                            probes[childPlan.SourceWebId] = Blocked(childPlan, "TargetParentMappingUnavailable", "The target parent Web mapping is unavailable.");
                            continue;
                        }
                        if (!parentProbe.IsAdmitted)
                        {
                            probes[childPlan.SourceWebId] = Blocked(childPlan, "TargetParentBlocked", "The target parent Web is blocked.");
                            continue;
                        }
                        if (!parentProbe.Exists)
                        {
                            probes[childPlan.SourceWebId] = PlannedCreate(childPlan, site.Id, parentProbe.TargetWebId);
                            continue;
                        }
                        probes[childPlan.SourceWebId] = InspectChild(
                            anchorContext,
                            childPlan,
                            site.Id,
                            parentProbe.TargetWebId.Value,
                            approvedHostId,
                            approvedHostUrl);
                    }
                    probe.Webs = plan.Webs.Select(value => probes[value.SourceWebId]).ToList();
                    return probe;
                }
            }
            catch (Exception exception) when (exception is ServerException || exception is ClientRequestException)
            {
                probe.Issues.Add(Issue("TargetSiteUnavailable", "target-site:" + plan.TargetSiteCollectionUrl,
                    "The planned target Site Collection could not be inspected: " + exception.Message));
                return probe;
            }
        }

        private static TopologyWebTargetProbe InspectChild(
            ClientContext anchorContext,
            WebMappingPlan plan,
            Guid targetSiteId,
            Guid targetParentWebId,
            Guid approvedHostId,
            string approvedHostUrl)
        {
            using (var context = anchorContext.Clone(plan.TargetParentWebUrl))
            {
                var parent = context.Web;
                context.Load(parent, value => value.Id, value => value.Url);
                context.Load(parent.Webs, values => values.Include(
                    value => value.Id,
                    value => value.Url,
                    value => value.ServerRelativeUrl,
                    value => value.Title,
                    value => value.Description,
                    value => value.WebTemplate,
                    value => value.Configuration,
                    value => value.AllProperties));
                context.ExecuteQueryRetry();
                if (parent.Id != targetParentWebId || !UrlEquals(parent.Url, plan.TargetParentWebUrl))
                {
                    return Blocked(plan, "TargetParentIdentityMismatch", "The observed target parent Web differs from the topology plan.");
                }

                var candidate = parent.Webs.AsEnumerable().SingleOrDefault(value =>
                    string.Equals(NormalizePath(value.ServerRelativeUrl), NormalizePath(plan.TargetServerRelativeUrl), StringComparison.OrdinalIgnoreCase));
                if (candidate == null)
                {
                    return PlannedCreate(plan, targetSiteId, targetParentWebId);
                }

                var originalIdentifier = Property(candidate.AllProperties, TopologyPlanner.WebOriginalIdentifierPropertyName);
                var mappingDigest = Property(candidate.AllProperties, TopologyPlanner.WebPlanDigestPropertyName);
                var expectedDigest = TopologyPlanner.ComputeWebMappingDigest(plan);
                var exactShape = string.Equals(candidate.Title, plan.TargetTitle, StringComparison.Ordinal)
                    && TemplateMatches(candidate.WebTemplate, candidate.Configuration, plan.TargetTemplate, plan.TargetConfiguration);
                var disposition = TopologyMaterializationDisposition.Block;
                if (candidate.Id == approvedHostId && UrlEquals(candidate.Url, approvedHostUrl))
                {
                    disposition = TopologyMaterializationDisposition.ReuseApprovedHost;
                }
                else if (exactShape
                    && string.Equals(originalIdentifier, plan.OriginalIdentifier, StringComparison.Ordinal)
                    && string.Equals(mappingDigest, expectedDigest, StringComparison.OrdinalIgnoreCase))
                {
                    disposition = TopologyMaterializationDisposition.ReuseOwned;
                }
                else if (exactShape
                    && string.IsNullOrWhiteSpace(originalIdentifier)
                    && string.IsNullOrWhiteSpace(mappingDigest)
                    && string.Equals(candidate.Description, InterruptedCreateDescription(plan), StringComparison.Ordinal))
                {
                    disposition = TopologyMaterializationDisposition.RecoverInterruptedCreate;
                }

                var result = new TopologyWebTargetProbe
                {
                    SourceSiteId = plan.SourceSiteId,
                    SourceWebId = plan.SourceWebId,
                    TargetWebUrl = plan.TargetWebUrl,
                    Exists = true,
                    TargetSiteId = targetSiteId,
                    TargetWebId = candidate.Id,
                    TargetParentWebId = targetParentWebId,
                    ExistingTitle = candidate.Title,
                    ExistingTemplate = candidate.WebTemplate,
                    ExistingConfiguration = candidate.Configuration,
                    ExistingOriginalIdentifier = originalIdentifier,
                    ExistingPlanDigest = mappingDigest,
                    Disposition = disposition
                };
                if (disposition == TopologyMaterializationDisposition.Block)
                {
                    result.Issues.Add(Issue("TargetWebOwnershipCollision", "target-web:" + plan.TargetWebUrl,
                        "The target child-Web path is occupied without approved-host identity or exact migration provenance, mapping digest, title, and template."));
                }
                return result;
            }
        }

        internal static string InterruptedCreateDescription(WebMappingPlan plan)
        {
            return "PnP migration mapping for " + plan.OriginalIdentifier;
        }

        private static TopologyWebTargetProbe PlannedCreate(WebMappingPlan plan, Guid targetSiteId, Guid? parentWebId)
        {
            return new TopologyWebTargetProbe
            {
                SourceSiteId = plan.SourceSiteId,
                SourceWebId = plan.SourceWebId,
                TargetWebUrl = plan.TargetWebUrl,
                Exists = false,
                TargetSiteId = targetSiteId,
                TargetParentWebId = parentWebId,
                Disposition = TopologyMaterializationDisposition.CreateOwned
            };
        }

        private static TopologyWebTargetProbe Blocked(WebMappingPlan plan, string code, string message)
        {
            var result = new TopologyWebTargetProbe
            {
                SourceSiteId = plan.SourceSiteId,
                SourceWebId = plan.SourceWebId,
                TargetWebUrl = plan.TargetWebUrl,
                Disposition = TopologyMaterializationDisposition.Block
            };
            result.Issues.Add(Issue(code, "target-web:" + plan.TargetWebUrl, message));
            return result;
        }

        private static bool TemplateMatches(string observedTemplate, int observedConfiguration, string expectedTemplate, int expectedConfiguration)
        {
            var parts = (expectedTemplate ?? string.Empty).Split('#');
            var template = parts[0];
            var configuration = parts.Length > 1
                ? int.Parse(parts[1], CultureInfo.InvariantCulture)
                : expectedConfiguration;
            return string.Equals(observedTemplate, template, StringComparison.OrdinalIgnoreCase)
                && observedConfiguration == configuration;
        }

        private static string Property(PropertyValues values, string name)
        {
            object value;
            return values != null && values.FieldValues.TryGetValue(name, out value) ? Convert.ToString(value) : null;
        }

        private static int Depth(string path)
        {
            return NormalizePath(path).Count(value => value == '/');
        }

        private static bool UrlEquals(string left, string right)
        {
            return string.Equals(NormalizeAbsolute(left), NormalizeAbsolute(right), StringComparison.OrdinalIgnoreCase);
        }

        private static string NormalizeAbsolute(string value)
        {
            return new Uri(value).AbsoluteUri.TrimEnd('/');
        }

        private static string NormalizePath(string value)
        {
            return Uri.UnescapeDataString(value ?? string.Empty).TrimEnd('/');
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
