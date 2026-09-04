using PnP.Framework.Migration.Diagnostics;
using PnP.Framework.Migration.Packaging;
using PnP.Framework.Migration.Taxonomy;
using PnP.Framework.Migration.Topology;
using System;
using System.Collections.Generic;
using System.Linq;

namespace PnP.Framework.Migration.Schema.ContentTypes
{
    public static class ContentTypeClosurePlanner
    {
        public static ContentTypeClosurePlanBuildResult Create(
            IEnumerable<ContentTypeSchemaSnapshot> snapshots,
            TopologyPlan topology,
            IEnumerable<TaxonomyTargetMapping> taxonomyMappings)
        {
            var result = new ContentTypeClosurePlanBuildResult();
            var mappings = topology.SiteCollections.SelectMany(value => value.Webs).ToArray();
            foreach (var snapshot in (snapshots ?? Enumerable.Empty<ContentTypeSchemaSnapshot>())
                         .OrderBy(value => value.ContentTypeId == null ? 0 : value.ContentTypeId.Length)
                         .ThenBy(value => value.ContentTypeId, StringComparer.OrdinalIgnoreCase))
            {
                var sourceScope = ScopePath(snapshot.SourceScope, snapshot.SourceWebUrl);
                var owner = mappings.SingleOrDefault(value => string.Equals(
                    NormalizePath(value.SourceServerRelativeUrl),
                    sourceScope,
                    StringComparison.OrdinalIgnoreCase));
                if (owner == null)
                {
                    result.Issues.Add(Issue("ContentTypeOwnerTopologyUnavailable", snapshot.ContentTypeId,
                        "No captured source Web mapping owns content type scope '" + sourceScope + "'."));
                    continue;
                }

                ContentTypeMaterializationPlan schema;
                try
                {
                    schema = ContentTypeRuntimeCatalog.IsTargetRuntime(snapshot.ContentTypeId)
                        && ContentTypeSchemaPlanner.TryCreateTargetRuntimeRequirement(snapshot, out var targetRuntimeRequirement)
                            ? targetRuntimeRequirement
                            : ContentTypeSchemaPlanner.CreateRequiredClosure(snapshot, taxonomyMappings);
                }
                catch (ArgumentException exception)
                {
                    result.Issues.Add(Issue("ContentTypeSchemaUnavailable", snapshot.ContentTypeId, exception.Message));
                    continue;
                }
                foreach (var blockedField in schema.Fields.Where(value => value.Disposition == Schema.Fields.FieldSchemaMaterializationDisposition.Block))
                {
                    result.Issues.Add(Issue("ContentTypeFieldMaterializerUnavailable", snapshot.ContentTypeId + ":" + blockedField.FieldId.ToString("D"), blockedField.Reason));
                }
                var node = new ContentTypeClosureNodePlan
                {
                    SourceOwnerWebId = owner.SourceWebId,
                    SourceOwnerWebUrl = owner.SourceWebUrl,
                    TargetOwnerWebUrl = owner.TargetWebUrl,
                    Schema = schema
                };
                node.PlanDigest = ComputeDigest(node);
                result.Nodes.Add(node);
            }
            result.Issues = result.Issues.OrderBy(value => value.Code, StringComparer.Ordinal)
                .ThenBy(value => value.Subject, StringComparer.Ordinal).ToList();
            return result;
        }

        public static string ComputeDigest(ContentTypeClosureNodePlan plan)
        {
            var digest = plan.PlanDigest;
            var probe = plan.TargetProbe;
            var admission = plan.TargetAdmission;
            var deferred = plan.DeferredUntilTopologyMaterialization;
            plan.PlanDigest = null;
            plan.TargetProbe = null;
            plan.TargetAdmission = null;
            plan.DeferredUntilTopologyMaterialization = false;
            try
            {
                return MigrationDigest.ComputeSha256(MigrationContractSerializer.SerializeCanonical(plan));
            }
            finally
            {
                plan.PlanDigest = digest;
                plan.TargetProbe = probe;
                plan.TargetAdmission = admission;
                plan.DeferredUntilTopologyMaterialization = deferred;
            }
        }

        private static string ScopePath(string scope, string sourceWebUrl)
        {
            if (string.IsNullOrWhiteSpace(scope))
            {
                return NormalizePath(new Uri(sourceWebUrl).AbsolutePath);
            }
            Uri absolute;
            return Uri.TryCreate(scope, UriKind.Absolute, out absolute)
                ? NormalizePath(absolute.AbsolutePath)
                : NormalizePath(scope);
        }

        private static string NormalizePath(string value)
        {
            var path = Uri.UnescapeDataString(value ?? string.Empty).Replace('\\', '/').TrimEnd('/');
            return path.Length == 0 ? "/" : path;
        }

        private static MigrationIssue Issue(string code, string subject, string message)
        {
            return new MigrationIssue
            {
                Code = code,
                Severity = MigrationIssueSeverity.Blocker,
                Subject = "content-type:" + subject,
                Ingredient = "ContentTypeClosure",
                Message = message
            };
        }
    }
}
