using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace PnP.Framework.Migration.Topology.Execution
{
    /// <summary>
    /// Produces an ancestor-closed topology transaction plan from the sealed
    /// approved plan. Mapping identities are not recalculated or remapped.
    /// </summary>
    internal static class TopologyExecutionPlanProjector
    {
        public static TopologyPlan Project(
            TopologyPlan approvedPlan,
            ISet<string> executableWebIngredientIds,
            Func<Guid, Guid, string> ingredientId)
        {
            if (approvedPlan == null || executableWebIngredientIds == null || executableWebIngredientIds.Count == 0)
            {
                return null;
            }
            if (ingredientId == null)
            {
                throw new ArgumentNullException(nameof(ingredientId));
            }

            var result = new TopologyPlan { SchemaVersion = approvedPlan.SchemaVersion };
            foreach (var site in approvedPlan.SiteCollections ?? Array.Empty<SiteCollectionMappingPlan>())
            {
                var selected = (site.Webs ?? Array.Empty<WebMappingPlan>())
                    .Where(value => value != null
                        && executableWebIngredientIds.Contains(ingredientId(value.SourceSiteId, value.SourceWebId)))
                    .ToList();
                if (selected.Count == 0)
                {
                    continue;
                }

                var selectedIds = new HashSet<Guid>(selected.Select(value => value.SourceWebId));
                var missingParent = selected.FirstOrDefault(value => value.SourceParentWebId.HasValue
                    && !selectedIds.Contains(value.SourceParentWebId.Value));
                if (missingParent != null)
                {
                    throw new InvalidDataException(
                        "The executable topology frontier is not ancestor-closed for source Web "
                        + missingParent.SourceWebId.ToString("D") + ".");
                }

                result.SiteCollections.Add(Clone(site, selected));
            }
            if (result.SiteCollections.Count == 0)
            {
                return null;
            }
            result.PlanDigest = TopologyPlanner.ComputeDigest(result);
            TopologyPlanValidator.Validate(result);
            return result;
        }

        private static SiteCollectionMappingPlan Clone(
            SiteCollectionMappingPlan source,
            IList<WebMappingPlan> webs)
        {
            return new SiteCollectionMappingPlan
            {
                SourceSiteId = source.SourceSiteId,
                SourceSiteCollectionUrl = source.SourceSiteCollectionUrl,
                TargetMode = source.TargetMode,
                PreferredTargetSiteCollectionUrl = source.PreferredTargetSiteCollectionUrl,
                TargetSiteCollectionUrl = source.TargetSiteCollectionUrl,
                TargetSiteCollisionResolved = source.TargetSiteCollisionResolved,
                TargetSiteResolutionReason = source.TargetSiteResolutionReason,
                ExpectedTargetSiteId = source.ExpectedTargetSiteId,
                TargetTitle = source.TargetTitle,
                TargetOwner = source.TargetOwner,
                TargetTemplate = source.TargetTemplate,
                TargetLanguage = source.TargetLanguage,
                TargetTimeZone = source.TargetTimeZone,
                OriginalIdentifier = source.OriginalIdentifier,
                Webs = webs.ToList()
            };
        }
    }
}
