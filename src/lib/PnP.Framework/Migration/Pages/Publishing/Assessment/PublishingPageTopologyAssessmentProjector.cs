using PnP.Framework.Migration.Evidence;
using PnP.Framework.Migration.Pages.Assessment;
using PnP.Framework.Migration.Pages.Ingredients;
using PnP.Framework.Migration.Pages.Publishing.Ingredients;
using PnP.Framework.Migration.Topology;
using System;
using System.Collections.Generic;
using System.Linq;

namespace PnP.Framework.Migration.Pages.Publishing.Assessment
{
    internal static class PublishingPageTopologyAssessmentProjector
    {
        public static void Project(
            PublishingPageAssessmentContext context,
            PublishingPageAssessmentAccumulator assessments)
        {
            var sourceTopology = context.Snapshot.SourceTopology;
            var mappings = (context.TargetSite?.Webs ?? Array.Empty<WebMappingPlan>())
                .Where(value => value != null)
                .GroupBy(value => Key(value.SourceSiteId, value.SourceWebId), StringComparer.Ordinal)
                .ToDictionary(group => group.Key, group => group.First(), StringComparer.Ordinal);
            var assessedWebIds = new HashSet<Guid>();
            foreach (var sourceWeb in (sourceTopology?.Webs ?? Array.Empty<SourceWebSnapshot>())
                         .Where(value => value != null))
            {
                assessedWebIds.Add(sourceWeb.WebId);
                AddAssessment(
                    assessments,
                    mappings,
                    sourceWeb.SiteId,
                    sourceWeb.WebId,
                    sourceWeb.Availability is EvidenceAvailability.Unavailable
                        or EvidenceAvailability.Conflict);
            }

            var source = context.Snapshot.Source;
            if (source != null
                && source.SiteId != Guid.Empty
                && source.WebId != Guid.Empty
                && !assessedWebIds.Contains(source.WebId))
            {
                AddAssessment(
                    assessments,
                    mappings,
                    source.SiteId,
                    source.WebId,
                    sourceUnavailable: false);
            }
        }

        private static void AddAssessment(
            PublishingPageAssessmentAccumulator assessments,
            IReadOnlyDictionary<string, WebMappingPlan> mappings,
            Guid sourceSiteId,
            Guid sourceWebId,
            bool sourceUnavailable)
        {
            mappings.TryGetValue(Key(sourceSiteId, sourceWebId), out var mapping);
            var blocked = sourceUnavailable || mapping == null;
            assessments.Add(
                PublishingPageIngredientIds.Web(sourceSiteId, sourceWebId),
                blocked
                    ? PageIngredientAssessmentState.KnownGap
                    : PageIngredientAssessmentState.TargetInspectionRequired,
                blocked
                    ? sourceUnavailable ? IngredientCapability.Missing : IngredientCapability.Incompatible
                    : IngredientCapability.Available,
                blocked ? IngredientDisposition.Defer : IngredientDisposition.Preserve,
                blocked ? "none" : "create-or-reuse-at-exact-relative-path",
                "policy.topology.web.exact-relative-path",
                sourceUnavailable
                    ? "The source Web topology evidence is unavailable or conflicting."
                    : mapping == null
                        ? "The captured source Web has no entry in the reviewed exact-relative-path topology plan."
                        : "Preserve the captured Site/Web hierarchy and relative URL segments; target inspection selects create, owned reuse, recovery, or a suffix only at an observed foreign collision node.",
                mapping?.TargetWebUrl,
                blocked
                    ? sourceUnavailable ? "SourceWebTopologyEvidenceUnavailable" : "TargetWebTopologyMappingUnavailable"
                    : null,
                blocked ? null : $"The topology receipt maps source Web '{sourceWebId:D}' to '{mapping.TargetWebUrl}'.",
                blocked ? null : "Fresh readback verifies target parent, path, template, identity, and ownership markers.");
        }

        private static string Key(Guid siteId, Guid webId)
        {
            return siteId.ToString("D") + "/" + webId.ToString("D");
        }
    }
}
