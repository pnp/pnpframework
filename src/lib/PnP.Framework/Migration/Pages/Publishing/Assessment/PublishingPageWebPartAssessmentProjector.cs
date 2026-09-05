using PnP.Framework.Migration.Lists.Planning;
using PnP.Framework.Migration.Pages.Assessment;
using PnP.Framework.Migration.Pages.ClassicWebParts.Bindings;
using PnP.Framework.Migration.Pages.ClassicWebParts.Planning;
using PnP.Framework.Migration.Pages.Ingredients;
using PnP.Framework.Migration.Pages.Publishing.Ingredients;
using System;
using System.Collections.Generic;
using System.Linq;

namespace PnP.Framework.Migration.Pages.Publishing.Assessment
{
    internal static class PublishingPageWebPartAssessmentProjector
    {
        public static void Project(
            PublishingPageAssessmentContext context,
            PublishingPageAssessmentAccumulator assessments)
        {
            var bindings = context.Snapshot.ListWebPartBindings
                .Where(value => value != null)
                .GroupBy(value => value.SourceWebPartId)
                .ToDictionary(group => group.Key, group => group.First());
            var listPlans = (context.ListPlan?.Lists ?? Array.Empty<ListMaterializationPlan>())
                .Where(value => value != null)
                .GroupBy(value => value.SourceListId)
                .ToDictionary(group => group.Key, group => group.First());
            foreach (var webPart in context.Snapshot.WebParts.Where(value => value != null))
            {
                var ingredientId = PublishingPageIngredientIds.WebPart(webPart.Id);
                var portabilityBlocker = ClassicWebPartReplayCapabilityPolicy.GetBlocker(webPart.ExportXml);
                if (!string.IsNullOrWhiteSpace(portabilityBlocker))
                {
                    assessments.Add(
                        ingredientId,
                        PageIngredientAssessmentState.KnownGap,
                        IngredientCapability.Incompatible,
                        IngredientDisposition.Defer,
                        "none",
                        "policy.webpart.classic",
                        portabilityBlocker,
                        mitigationCode: "WebPartCapabilityUnavailable");
                    continue;
                }

                if (!bindings.TryGetValue(webPart.Id, out var binding))
                {
                    assessments.Add(
                        ingredientId,
                        PageIngredientAssessmentState.TargetInspectionRequired,
                        IngredientCapability.Available,
                        IngredientDisposition.Preserve,
                        "copy-captured-export",
                        "policy.webpart.classic",
                        "Copy the portable captured shared Web Part export after deterministic text rewrites.",
                        context.TargetPageServerRelativeUrl,
                        null,
                        $"Fresh target inspection and readback verify Web Part '{webPart.Id:D}' placement, type, properties, and zone.");
                    continue;
                }

                if (!listPlans.TryGetValue(binding.SourceListId, out var listPlan))
                {
                    assessments.Add(
                        ingredientId,
                        PageIngredientAssessmentState.KnownGap,
                        IngredientCapability.Missing,
                        IngredientDisposition.Defer,
                        "none",
                        "policy.webpart.classic-list-binding",
                        "The bound source List has no source-authoritative target path plan.",
                        mitigationCode: "ListWebPartBindingPlanUnavailable");
                    continue;
                }

                var viewPlan = binding.SourceViewId.HasValue
                    ? listPlan.Views.FirstOrDefault(value => value.SourceViewId == binding.SourceViewId.Value)
                    : null;
                if (!binding.SourceViewId.HasValue
                    || viewPlan == null
                    || viewPlan.Disposition is ListViewMaterializationDisposition.Block
                        or ListViewMaterializationDisposition.SkipPersonal)
                {
                    assessments.Add(
                        ingredientId,
                        PageIngredientAssessmentState.KnownGap,
                        IngredientCapability.Incompatible,
                        IngredientDisposition.Defer,
                        "none",
                        "policy.webpart.classic-list-binding",
                        "The captured List Web Part has no executable source View identity and candidate target View action.",
                        mitigationCode: "ListWebPartViewMappingUnavailable");
                    continue;
                }

                assessments.Add(
                    ingredientId,
                    PageIngredientAssessmentState.TargetInspectionRequired,
                    IngredientCapability.Available,
                    IngredientDisposition.Preserve,
                    "rewrite-list-and-view-binding",
                    "policy.webpart.classic-list-binding",
                    "After List materialization, resolve target runtime Web/List/View IDs and rewrite only the sealed binding properties.",
                    listPlan.TargetRootFolderServerRelativeUrl,
                    null,
                    $"Fresh readback verifies Web Part '{webPart.Id:D}' placement and the exact approved target List/View binding.");
            }
        }
    }
}
