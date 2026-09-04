using PnP.Framework.Migration.Pages.Assessment;
using PnP.Framework.Migration.Pages.Ingredients;
using PnP.Framework.Migration.Pages.Publishing.Ingredients;
using PnP.Framework.Migration.Pages.References;
using System;
using System.Collections.Generic;
using System.Linq;

namespace PnP.Framework.Migration.Pages.Publishing.Assessment
{
    internal static class PublishingPageReferenceAssessmentProjector
    {
        public static void Project(
            PublishingPageAssessmentContext context,
            PublishingPageAssessmentAccumulator assessments)
        {
            var actions = context.ReferenceActions
                .Where(value => value != null)
                .GroupBy(value => value.SnapshotDependencyId, StringComparer.Ordinal)
                .ToDictionary(group => group.Key, group => group.First(), StringComparer.Ordinal);
            foreach (var reference in context.Snapshot.Dependencies.Where(value => value != null))
            {
                actions.TryGetValue(reference.Id, out var action);
                var ingredientId = PublishingPageIngredientIds.Reference(reference.Id);
                if (action == null)
                {
                    assessments.Add(
                        ingredientId,
                        PageIngredientAssessmentState.KnownGap,
                        IngredientCapability.Missing,
                        IngredientDisposition.Defer,
                        "none",
                        "policy.reference.page",
                        context.ReferencePlanningFailure ?? "No source-authoritative reference action was produced.",
                        mitigationCode: "PageReferenceActionUnavailable");
                    continue;
                }

                switch (action.Disposition)
                {
                    case PageReferenceDisposition.PreserveExternal:
                        assessments.Add(
                            ingredientId,
                            PageIngredientAssessmentState.Determined,
                            IngredientCapability.Available,
                            IngredientDisposition.Preserve,
                            "reuse-external-reference",
                            "policy.reference.page",
                            Reason(action, "Preserve the external reference exactly as captured."),
                            action.TargetAbsoluteUrl ?? reference.SourceAbsoluteUrl,
                            null,
                            "Stored content preserves the reviewed external reference.");
                        break;
                    case PageReferenceDisposition.RewriteToTarget:
                        assessments.Add(
                            ingredientId,
                            PageIngredientAssessmentState.Determined,
                            IngredientCapability.Available,
                            IngredientDisposition.Transform,
                            "rewrite-reference",
                            "policy.reference.page",
                            Reason(action, "Rewrite the same-tenant source reference to the exact mapped target path."),
                            action.TargetServerRelativeUrl ?? action.TargetAbsoluteUrl,
                            null,
                            "Stored content contains the deterministic mapped target reference.");
                        break;
                    case PageReferenceDisposition.MaterializeAtTarget:
                        assessments.Add(
                            ingredientId,
                            PageIngredientAssessmentState.TargetInspectionRequired,
                            IngredientCapability.Available,
                            IngredientDisposition.Preserve,
                            "copy-exact-bytes-create-only",
                            "policy.reference.page",
                            Reason(action, "Materialize the captured exact resource bytes at the mapped target path."),
                            action.TargetServerRelativeUrl ?? action.TargetAbsoluteUrl,
                            null,
                            $"Fresh readback verifies resource bytes with SHA-256 '{reference.ContentSha256}'.");
                        break;
                    case PageReferenceDisposition.Delegate:
                        assessments.Add(
                            ingredientId,
                            PageIngredientAssessmentState.Determined,
                            IngredientCapability.Unknown,
                            IngredientDisposition.Delegate,
                            "retain-snapshot",
                            "policy.reference.page",
                            Reason(action, "Retain the complete reference evidence for a separately reviewed workflow."),
                            verificationAssertions: "The reference evidence remains digest-verifiable in the source snapshot.");
                        break;
                    default:
                        assessments.Add(
                            ingredientId,
                            PageIngredientAssessmentState.KnownGap,
                            IngredientCapability.Incompatible,
                            IngredientDisposition.Defer,
                            "none",
                            "policy.reference.page",
                            Reason(action, "The captured reference has no fidelity-preserving target action."),
                            mitigationCode: "PageReferenceMaterializationUnavailable");
                        break;
                }
            }
        }

        private static string Reason(PageReferenceAction action, string fallback)
        {
            var diagnostics = (action.Diagnostics ?? Array.Empty<string>())
                .Where(value => !string.IsNullOrWhiteSpace(value))
                .Distinct(StringComparer.Ordinal)
                .ToArray();
            return diagnostics.Length == 0 ? fallback : string.Join("; ", diagnostics);
        }
    }
}
