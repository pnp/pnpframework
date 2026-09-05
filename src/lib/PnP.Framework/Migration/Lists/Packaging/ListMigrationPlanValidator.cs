using PnP.Framework.Migration.Lists.Capture;
using PnP.Framework.Migration.Lists.Planning;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using PnP.Framework.Migration.Schema.ContentTypes;
using PnP.Framework.Migration.Schema.ContentTypes.Packaging;
using PnP.Framework.Migration.Features;

namespace PnP.Framework.Migration.Lists.Packaging
{
    internal static class ListMigrationPlanValidator
    {
        public static void Validate(IEnumerable<ListDependencySnapshot> snapshots, ListMigrationPlanSet plan)
        {
            var sources = (snapshots ?? Enumerable.Empty<ListDependencySnapshot>()).ToArray();
            if (sources.Length == 0 && plan == null)
            {
                return;
            }
            if (plan == null || plan.Lists == null || plan.OrderedSourceListIds == null || plan.Issues == null)
            {
                throw new InvalidDataException("Captured List dependencies require a complete List migration plan.");
            }
            var sourceIds = new HashSet<Guid>(sources.Select(value => value.SourceListId));
            var sourceById = sources.ToDictionary(value => value.SourceListId);
            var plannedIds = new HashSet<Guid>(plan.Lists.Select(value => value == null ? Guid.Empty : value.SourceListId));
            if (plan.Lists.Any(value => value == null) || sourceIds.Count != plannedIds.Count || !sourceIds.SetEquals(plannedIds))
            {
                throw new InvalidDataException("The List migration plan must contain exactly one plan for every captured List dependency.");
            }
            foreach (var list in plan.Lists)
            {
                if (list.Fields == null || list.Views == null || list.ViewRenderingResources == null
                    || list.SiteContentTypes == null || list.RequiredFeatures == null || list.Issues == null
                    || string.IsNullOrWhiteSpace(list.OriginalIdentifier) || string.IsNullOrWhiteSpace(list.TargetSiteCollectionUrl)
                    || !string.Equals(ListMigrationPlanFactory.ComputePlanDigest(list), list.PlanDigest, StringComparison.OrdinalIgnoreCase))
                {
                    throw new InvalidDataException("A List materialization plan is incomplete or its semantic digest differs: " + list.SourceListId.ToString("D"));
                }
                if (list.Disposition != ListMaterializationDisposition.Block && (list.TargetProbe == null || !list.TargetProbe.IsAdmitted))
                {
                    throw new InvalidDataException("An executable List plan has no admitted target probe: " + list.SourceListId.ToString("D"));
                }
                foreach (var contentType in list.SiteContentTypes)
                {
                    if (contentType == null || contentType.Schema == null
                        || !string.Equals(ContentTypeClosurePlanner.ComputeDigest(contentType), contentType.PlanDigest, StringComparison.OrdinalIgnoreCase))
                    {
                        throw new InvalidDataException("A site content type closure node is incomplete or its digest differs.");
                    }
                    ContentTypeSchemaContractValidator.ValidatePlan(contentType.Schema);
                    if (list.Disposition != ListMaterializationDisposition.Block
                        && !contentType.DeferredUntilTopologyMaterialization
                        && (contentType.TargetAdmission == null || !contentType.TargetAdmission.IsEligible))
                    {
                        throw new InvalidDataException("An executable site content type plan has no admitted target analysis: " + contentType.Schema.ContentTypeId);
                    }
                }
                ValidateViewRenderingResources(sourceById[list.SourceListId], list);
                ValidateFeatures(sourceById[list.SourceListId], list);
            }
            if (!string.Equals(ListMigrationPlanFactory.ComputeSetDigest(plan), plan.PlanDigest, StringComparison.OrdinalIgnoreCase))
            {
                throw new InvalidDataException("The List migration plan-set digest differs from its sealed content.");
            }
        }

        private static void ValidateViewRenderingResources(ListDependencySnapshot source, ListMaterializationPlan list)
        {
            var expected = source.ViewRenderingResources
                .ToDictionary(value => value.Id, StringComparer.Ordinal);
            var actual = list.ViewRenderingResources
                .ToDictionary(value => value == null ? string.Empty : value.SourceResourceId, StringComparer.Ordinal);
            if (actual.ContainsKey(string.Empty)
                || expected.Count != actual.Count
                || !new HashSet<string>(expected.Keys, StringComparer.Ordinal).SetEquals(actual.Keys))
            {
                throw new InvalidDataException("The View rendering-resource plan does not exactly cover the captured resource inventory: "
                    + list.SourceListId.ToString("D") + ".");
            }
            foreach (var pair in expected)
            {
                var planned = actual[pair.Key];
                if (!string.Equals(planned.SourceAbsoluteUrl, pair.Value.SourceAbsoluteUrl, StringComparison.OrdinalIgnoreCase)
                    || !string.Equals(planned.SourceServerRelativeUrl, pair.Value.SourceServerRelativeUrl, StringComparison.OrdinalIgnoreCase)
                    || !string.Equals(planned.SourceArtifact?.Sha256, pair.Value.Artifact?.Sha256, StringComparison.OrdinalIgnoreCase))
                {
                    throw new InvalidDataException("A View rendering-resource plan differs from its sealed source evidence: " + pair.Key + ".");
                }
            }
        }

        private static void ValidateFeatures(ListDependencySnapshot source, ListMaterializationPlan list)
        {
            var expected = ContentTypeRuntimeCatalog.CreateFeatureRequirements(
                source.ContentTypes.Select(value => value.ParentId),
                source.SiteContentTypes,
                list.TargetSiteCollectionUrl).ToDictionary(value => value.FeatureId);
            var actual = list.RequiredFeatures.ToDictionary(value => value == null ? Guid.Empty : value.FeatureId);
            if (actual.ContainsKey(Guid.Empty) || expected.Count != actual.Count || !new HashSet<Guid>(expected.Keys).SetEquals(actual.Keys))
            {
                throw new InvalidDataException("The List platform-feature plan does not exactly cover its conditional target-runtime content types: "
                    + list.SourceListId.ToString("D") + ".");
            }
            foreach (var pair in expected)
            {
                var observed = actual[pair.Key];
                var semanticMatch = observed.Scope == pair.Value.Scope
                    && observed.DependencyOrder == pair.Value.DependencyOrder
                    && observed.Disposition == pair.Value.Disposition
                    && string.Equals(observed.Name, pair.Value.Name, StringComparison.Ordinal)
                    && string.Equals(observed.TargetWebUrl, pair.Value.TargetWebUrl, StringComparison.OrdinalIgnoreCase)
                    && observed.DependsOnFeatureIds.SequenceEqual(pair.Value.DependsOnFeatureIds)
                    && observed.RequiredByContentTypeIds.SequenceEqual(pair.Value.RequiredByContentTypeIds, StringComparer.OrdinalIgnoreCase)
                    && observed.ExpectedContentTypeIds.SequenceEqual(pair.Value.ExpectedContentTypeIds, StringComparer.OrdinalIgnoreCase);
                if (!semanticMatch)
                {
                    throw new InvalidDataException("The platform-feature plan differs from the captured content-type requirement: "
                        + pair.Key.ToString("D") + ".");
                }
                if (list.Disposition != ListMaterializationDisposition.Block
                    && (observed.TargetProbe == null || !observed.TargetProbe.IsAdmitted))
                {
                    throw new InvalidDataException("An executable platform-feature plan has no admitted target probe: "
                        + pair.Key.ToString("D") + ".");
                }
            }
        }
    }
}
