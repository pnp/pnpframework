using Microsoft.SharePoint.Client;
using PnP.Framework.Migration.Diagnostics;
using System;
using System.Globalization;
using System.Linq;
using static PnP.Framework.Migration.Topology.TopologyTargetInspectionScope;

namespace PnP.Framework.Migration.Topology
{
    internal static class TopologyWebTargetInspector
    {
        public static TopologyWebTargetProbe Inspect(
            TopologyTargetInspectionScope scope,
            WebMappingPlan plan,
            Guid targetSiteId,
            Guid targetParentWebId)
        {
            return Inspect(scope, plan, targetSiteId, targetParentWebId, false);
        }

        public static TopologyWebTargetProbe InspectForPlanning(
            TopologyTargetInspectionScope scope,
            WebMappingPlan plan,
            Guid targetSiteId,
            Guid targetParentWebId)
        {
            return Inspect(scope, plan, targetSiteId, targetParentWebId, true);
        }

        private static TopologyWebTargetProbe Inspect(
            TopologyTargetInspectionScope scope,
            WebMappingPlan plan,
            Guid targetSiteId,
            Guid targetParentWebId,
            bool resolvePlanningCollision)
        {
            var approvedHost = scope.AnchorContext.Web;
            if (TargetUrl.Equals(plan.TargetWebUrl, scope.ApprovedHostUrl)
                && approvedHost.IsPropertyAvailable("Id")
                && approvedHost.IsPropertyAvailable("Url")
                && approvedHost.IsPropertyAvailable("Title")
                && approvedHost.IsPropertyAvailable("WebTemplate")
                && approvedHost.IsPropertyAvailable("Configuration")
                && approvedHost.IsObjectPropertyInstantiated("AllProperties")
                && approvedHost.Id == scope.ApprovedHostId
                && TargetUrl.Equals(approvedHost.Url, plan.TargetWebUrl))
            {
                return FromApprovedHost(approvedHost, plan, targetSiteId, targetParentWebId);
            }

            using (var context = scope.AnchorContext.Clone(plan.TargetParentWebUrl))
            {
                var parent = context.Web;
                context.Load(parent, value => value.Id, value => value.Url);
                if (resolvePlanningCollision)
                {
                    context.Load(parent.Webs, values => values.Include(
                        value => value.Id,
                        value => value.Url,
                        value => value.ServerRelativeUrl,
                        value => value.Title,
                        value => value.Description,
                        value => value.WebTemplate,
                        value => value.Configuration,
                        value => value.AllProperties));
                }
                else
                {
                    var targetServerRelativeUrl = plan.TargetServerRelativeUrl;
                    context.Load(parent.Webs, values => values
                        .Where(value => value.ServerRelativeUrl == targetServerRelativeUrl)
                        .Include(
                            value => value.Id,
                            value => value.Url,
                            value => value.ServerRelativeUrl,
                            value => value.Title,
                            value => value.Description,
                            value => value.WebTemplate,
                            value => value.Configuration,
                            value => value.AllProperties));
                }
                context.ExecuteQueryRetry();
                if (parent.Id != targetParentWebId || !TargetUrl.Equals(parent.Url, plan.TargetParentWebUrl))
                {
                    return Blocked(plan, "TargetParentIdentityMismatch", "The observed target parent Web differs from the topology plan.");
                }

                if (resolvePlanningCollision)
                {
                    var resolution = TopologyWebTargetPathResolver.Resolve(
                        plan,
                        parent.Webs.AsEnumerable().Select(value => new TopologyWebTargetInventoryItem
                        {
                            WebId = value.Id,
                            Url = value.Url,
                            ServerRelativeUrl = value.ServerRelativeUrl,
                            Title = value.Title,
                            Description = value.Description,
                            Template = value.WebTemplate,
                            Configuration = value.Configuration,
                            OriginalIdentifier = Property(value.AllProperties, TopologyPlanner.WebOriginalIdentifierPropertyName),
                            MappingDigest = Property(value.AllProperties, TopologyPlanner.WebPlanDigestPropertyName)
                        }));
                    if (resolution.ExistingTarget == null)
                    {
                        return PlannedCreate(
                            plan,
                            targetSiteId,
                            targetParentWebId,
                            resolution.TargetWebUrl,
                            resolution.TargetServerRelativeUrl,
                            resolution.CollisionResolved,
                            resolution.Reason);
                    }
                    return FromResolution(plan, targetSiteId, targetParentWebId, resolution);
                }

                var candidate = parent.Webs.AsEnumerable().SingleOrDefault(value =>
                    string.Equals(TargetUrl.NormalizePath(value.ServerRelativeUrl), TargetUrl.NormalizePath(plan.TargetServerRelativeUrl), StringComparison.OrdinalIgnoreCase));
                if (candidate == null)
                {
                    return PlannedCreate(plan, targetSiteId, targetParentWebId);
                }

                var originalIdentifier = Property(candidate.AllProperties, TopologyPlanner.WebOriginalIdentifierPropertyName);
                var mappingDigest = Property(candidate.AllProperties, TopologyPlanner.WebPlanDigestPropertyName);
                var exactShape = string.Equals(candidate.Title, plan.TargetTitle, StringComparison.Ordinal)
                    && TemplateMatches(candidate.WebTemplate, candidate.Configuration, plan.TargetTemplate, plan.TargetConfiguration);
                var disposition = ResolveDisposition(scope, plan, candidate, originalIdentifier, mappingDigest, exactShape);
                var result = new TopologyWebTargetProbe
                {
                    SourceSiteId = plan.SourceSiteId,
                    SourceWebId = plan.SourceWebId,
                    PreferredTargetWebUrl = plan.PreferredTargetWebUrl ?? plan.TargetWebUrl,
                    TargetWebUrl = plan.TargetWebUrl,
                    PreferredTargetServerRelativeUrl = plan.PreferredTargetServerRelativeUrl ?? plan.TargetServerRelativeUrl,
                    TargetServerRelativeUrl = plan.TargetServerRelativeUrl,
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

        private static TopologyWebTargetProbe FromApprovedHost(Web approvedHost, WebMappingPlan plan, Guid targetSiteId, Guid targetParentWebId)
        {
            return new TopologyWebTargetProbe
            {
                SourceSiteId = plan.SourceSiteId,
                SourceWebId = plan.SourceWebId,
                PreferredTargetWebUrl = plan.PreferredTargetWebUrl ?? plan.TargetWebUrl,
                TargetWebUrl = plan.TargetWebUrl,
                PreferredTargetServerRelativeUrl = plan.PreferredTargetServerRelativeUrl ?? plan.TargetServerRelativeUrl,
                TargetServerRelativeUrl = plan.TargetServerRelativeUrl,
                Exists = true,
                TargetSiteId = targetSiteId,
                TargetWebId = approvedHost.Id,
                TargetParentWebId = targetParentWebId,
                ExistingTitle = approvedHost.Title,
                ExistingTemplate = approvedHost.WebTemplate,
                ExistingConfiguration = approvedHost.Configuration,
                ExistingOriginalIdentifier = Property(approvedHost.AllProperties, TopologyPlanner.WebOriginalIdentifierPropertyName),
                ExistingPlanDigest = Property(approvedHost.AllProperties, TopologyPlanner.WebPlanDigestPropertyName),
                Disposition = TopologyMaterializationDisposition.ReuseApprovedHost
            };
        }

        private static TopologyMaterializationDisposition ResolveDisposition(
            TopologyTargetInspectionScope scope,
            WebMappingPlan plan,
            Web candidate,
            string originalIdentifier,
            string mappingDigest,
            bool exactShape)
        {
            if (candidate.Id == scope.ApprovedHostId && TargetUrl.Equals(candidate.Url, scope.ApprovedHostUrl))
            {
                return TopologyMaterializationDisposition.ReuseApprovedHost;
            }
            if (exactShape
                && string.Equals(originalIdentifier, plan.OriginalIdentifier, StringComparison.Ordinal)
                && string.Equals(mappingDigest, TopologyPlanner.ComputeWebMappingDigest(plan), StringComparison.OrdinalIgnoreCase))
            {
                return TopologyMaterializationDisposition.ReuseOwned;
            }
            if (exactShape
                && string.IsNullOrWhiteSpace(originalIdentifier)
                && string.IsNullOrWhiteSpace(mappingDigest)
                && string.Equals(candidate.Description, InterruptedCreateDescription(plan), StringComparison.Ordinal))
            {
                return TopologyMaterializationDisposition.RecoverInterruptedCreate;
            }
            return TopologyMaterializationDisposition.Block;
        }

        internal static string InterruptedCreateDescription(WebMappingPlan plan)
        {
            return "PnP migration mapping for " + plan.OriginalIdentifier;
        }

        internal static TopologyWebTargetProbe PlannedCreate(WebMappingPlan plan, Guid targetSiteId, Guid? parentWebId)
        {
            return PlannedCreate(
                plan,
                targetSiteId,
                parentWebId,
                plan.TargetWebUrl,
                plan.TargetServerRelativeUrl,
                false,
                null);
        }

        private static TopologyWebTargetProbe PlannedCreate(
            WebMappingPlan plan,
            Guid targetSiteId,
            Guid? parentWebId,
            string targetWebUrl,
            string targetServerRelativeUrl,
            bool collisionResolved,
            string collisionResolutionReason)
        {
            return new TopologyWebTargetProbe
            {
                SourceSiteId = plan.SourceSiteId,
                SourceWebId = plan.SourceWebId,
                PreferredTargetWebUrl = plan.PreferredTargetWebUrl ?? plan.TargetWebUrl,
                TargetWebUrl = targetWebUrl,
                PreferredTargetServerRelativeUrl = plan.PreferredTargetServerRelativeUrl ?? plan.TargetServerRelativeUrl,
                TargetServerRelativeUrl = targetServerRelativeUrl,
                CollisionResolved = collisionResolved,
                CollisionResolutionReason = collisionResolutionReason,
                Exists = false,
                TargetSiteId = targetSiteId,
                TargetParentWebId = parentWebId,
                Disposition = TopologyMaterializationDisposition.CreateOwned
            };
        }

        internal static TopologyWebTargetProbe Blocked(WebMappingPlan plan, string code, string message)
        {
            var result = new TopologyWebTargetProbe
            {
                SourceSiteId = plan.SourceSiteId,
                SourceWebId = plan.SourceWebId,
                PreferredTargetWebUrl = plan.PreferredTargetWebUrl ?? plan.TargetWebUrl,
                TargetWebUrl = plan.TargetWebUrl,
                PreferredTargetServerRelativeUrl = plan.PreferredTargetServerRelativeUrl ?? plan.TargetServerRelativeUrl,
                TargetServerRelativeUrl = plan.TargetServerRelativeUrl,
                Disposition = TopologyMaterializationDisposition.Block
            };
            result.Issues.Add(Issue(code, "target-web:" + plan.TargetWebUrl, message));
            return result;
        }

        internal static bool TemplateMatches(string observedTemplate, int observedConfiguration, string expectedTemplate, int expectedConfiguration)
        {
            var parts = (expectedTemplate ?? string.Empty).Split('#');
            var template = parts[0];
            var configuration = parts.Length > 1
                ? int.Parse(parts[1], CultureInfo.InvariantCulture)
                : expectedConfiguration;
            return string.Equals(observedTemplate, template, StringComparison.OrdinalIgnoreCase)
                && observedConfiguration == configuration;
        }

        private static TopologyWebTargetProbe FromResolution(
            WebMappingPlan plan,
            Guid targetSiteId,
            Guid targetParentWebId,
            TopologyWebTargetPathResolution resolution)
        {
            var existing = resolution.ExistingTarget;
            return new TopologyWebTargetProbe
            {
                SourceSiteId = plan.SourceSiteId,
                SourceWebId = plan.SourceWebId,
                PreferredTargetWebUrl = resolution.PreferredTargetWebUrl,
                TargetWebUrl = resolution.TargetWebUrl,
                PreferredTargetServerRelativeUrl = resolution.PreferredTargetServerRelativeUrl,
                TargetServerRelativeUrl = resolution.TargetServerRelativeUrl,
                CollisionResolved = resolution.CollisionResolved,
                CollisionResolutionReason = resolution.Reason,
                Exists = true,
                TargetSiteId = targetSiteId,
                TargetWebId = existing.WebId,
                TargetParentWebId = targetParentWebId,
                ExistingTitle = existing.Title,
                ExistingTemplate = existing.Template,
                ExistingConfiguration = existing.Configuration,
                ExistingOriginalIdentifier = existing.OriginalIdentifier,
                ExistingPlanDigest = existing.MappingDigest,
                Disposition = resolution.ExistingDisposition
            };
        }

        private static string Property(PropertyValues values, string name)
        {
            object value;
            return values != null && values.FieldValues.TryGetValue(name, out value) ? Convert.ToString(value) : null;
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
