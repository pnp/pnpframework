using PnP.Framework.Migration.Pages.Fields;
using PnP.Framework.Migration.Pages.Ingredients;
using PnP.Framework.Migration.Pages.Publishing.Capture;
using PnP.Framework.Migration.Pages.Publishing.Lifecycle;
using PnP.Framework.Migration.Pages.Publishing.Planning;
using PnP.Framework.Migration.Pages.Runtime;
using PnP.Framework.Migration.Taxonomy;
using System;
using System.Collections.Generic;
using System.Linq;

namespace PnP.Framework.Migration.Pages.Publishing.Ingredients
{
    internal static class PublishingPageCoreIngredientActionProjector
    {
        public static void Project(
            PublishingPageCaptureBundle snapshot,
            PublishingPageMigrationPlan plan,
            IDictionary<string, PageIngredientAction> actions)
        {
            AddRuntimeAndArtifact(snapshot, plan, actions);
            AddContentSecurityAndLifecycle(snapshot, plan, actions);
            AddFields(snapshot, plan, actions);
            AddTaxonomyRelationships(snapshot, plan, actions);
        }

        private static void AddRuntimeAndArtifact(
            PublishingPageCaptureBundle snapshot,
            PublishingPageMigrationPlan plan,
            IDictionary<string, PageIngredientAction> actions)
        {
            var publishingRuntime = string.Equals(snapshot.Runtime?.AdapterId, PageRuntimeAdapterIds.Publishing, StringComparison.Ordinal);
            PublishingPageIngredientActionFactory.Add(actions, PublishingPageIngredientActionFactory.Create(
                PublishingPageIngredientIds.Runtime,
                publishingRuntime ? IngredientCapability.Available : IngredientCapability.Incompatible,
                publishingRuntime ? IngredientDisposition.Preserve : IngredientDisposition.Block,
                publishingRuntime ? "reuse-target-runtime" : "none",
                "policy.runtime.publishing",
                publishingRuntime
                    ? "The Publishing CLR runtime selects the executable Publishing adapter."
                    : $"Runtime adapter '{snapshot.Runtime?.AdapterId ?? PageRuntimeAdapterIds.Unknown}' is not executable by the Publishing importer.",
                PageRuntimeAdapterIds.Publishing,
                "The target page resolves through the Publishing runtime without an error shell."));
            PublishingPageIngredientActionFactory.Add(actions, PublishingPageIngredientActionFactory.Create(
                PublishingPageIngredientIds.PageArtifact,
                publishingRuntime ? IngredientCapability.Available : IngredientCapability.Incompatible,
                publishingRuntime ? IngredientDisposition.Transform : IngredientDisposition.Block,
                publishingRuntime ? "create-target-page-shell" : "none",
                "policy.page-artifact.publishing",
                "Create a target Publishing page shell; retain the source ASPX bytes as immutable evidence rather than copying them as executable code.",
                plan.TargetPageServerRelativeUrl,
                "The exact target page path exists after mutation.",
                "The source ASPX artifact remains digest-verifiable in the package."));
        }

        private static void AddContentSecurityAndLifecycle(
            PublishingPageCaptureBundle snapshot,
            PublishingPageMigrationPlan plan,
            IDictionary<string, PageIngredientAction> actions)
        {
            var contentTransformed = plan.Replacements != null && plan.Replacements.Count > 0;
            PublishingPageIngredientActionFactory.Add(actions, PublishingPageIngredientActionFactory.Create(
                PublishingPageIngredientIds.PublishingContent,
                IngredientCapability.Available,
                contentTransformed ? IngredientDisposition.Transform : IngredientDisposition.Preserve,
                contentTransformed ? "copy-and-rewrite-approved-references" : "copy-exact-value",
                "policy.content.publishing",
                contentTransformed
                    ? "Write PublishingPageContent after applying the explicitly approved reference rewrites."
                    : "Write the captured PublishingPageContent without a text rewrite.",
                plan.TargetPageServerRelativeUrl + "#PublishingPageContent",
                $"The normalized target PublishingPageContent SHA-256 equals '{plan.ExpectedPublishingPageContentSha256}'."));

            var uniqueSecurity = snapshot.Security?.HasUniqueRoleAssignments == true;
            var securityBlocked = uniqueSecurity && plan.PlanningPolicy?.RequireInheritedPermissions == true;
            PublishingPageIngredientActionFactory.Add(actions, PublishingPageIngredientActionFactory.Create(
                PublishingPageIngredientIds.Security,
                securityBlocked ? IngredientCapability.Incompatible : IngredientCapability.Available,
                securityBlocked ? IngredientDisposition.Block : uniqueSecurity ? IngredientDisposition.Delegate : IngredientDisposition.Preserve,
                uniqueSecurity ? "retain-snapshot" : "reuse-target-inheritance",
                "policy.security.page",
                securityBlocked
                    ? "The source page has unique permissions, but this plan requires inherited target permissions."
                    : uniqueSecurity
                        ? "Unique permissions remain in the snapshot for a future principal-mapping workflow."
                        : "Reuse inherited target permissions.",
                plan.TargetPageServerRelativeUrl + "#security",
                uniqueSecurity
                    ? "Captured role assignments remain available in the source snapshot."
                    : "The target page inherits permissions."));

            var sourcePublished = string.Equals(snapshot.Lifecycle?.Level, "Published", StringComparison.OrdinalIgnoreCase);
            var lifecyclePreserved = sourcePublished == (plan.TargetLifecycle == PublishingPageTargetLifecycle.Published);
            PublishingPageIngredientActionFactory.Add(actions, PublishingPageIngredientActionFactory.Create(
                PublishingPageIngredientIds.Lifecycle,
                IngredientCapability.Available,
                lifecyclePreserved ? IngredientDisposition.Preserve : IngredientDisposition.Transform,
                plan.TargetLifecycle == PublishingPageTargetLifecycle.Published ? "publish-after-verification" : "check-in-as-draft",
                "policy.lifecycle.publishing",
                plan.LifecycleReason,
                plan.TargetPageServerRelativeUrl + "#lifecycle",
                $"The final target lifecycle is '{plan.TargetLifecycle}'."));
        }

        private static void AddFields(
            PublishingPageCaptureBundle snapshot,
            PublishingPageMigrationPlan plan,
            IDictionary<string, PageIngredientAction> actions)
        {
            var fields = snapshot.Fields.ToDictionary(value => value.InternalName, StringComparer.OrdinalIgnoreCase);
            foreach (var fieldAction in plan.FieldActions)
            {
                if (!fields.TryGetValue(fieldAction.SourceInternalName, out var field))
                {
                    continue;
                }

                var mapping = Map(fieldAction.Disposition, field.HasValue);
                PublishingPageIngredientActionFactory.Add(actions, PublishingPageIngredientActionFactory.Create(
                    PublishingPageIngredientIds.Field(fieldAction.SourceInternalName),
                    mapping.Capability,
                    mapping.Disposition,
                    mapping.Realization,
                    "policy.field." + fieldAction.Disposition.ToString().ToLowerInvariant(),
                    fieldAction.Reason,
                    fieldAction.TargetInternalName,
                    fieldAction.WillApply
                        ? $"The target field '{fieldAction.TargetInternalName}' round-trips the approved value."
                        : $"The reviewed '{fieldAction.Disposition}' decision remains sealed for field '{fieldAction.SourceInternalName}'."));
            }
        }

        private static (IngredientCapability Capability, IngredientDisposition Disposition, string Realization) Map(
            PageFieldDisposition disposition,
            bool hasValue)
        {
            switch (disposition)
            {
                case PageFieldDisposition.Apply:
                    return (IngredientCapability.Available, IngredientDisposition.Preserve, "map-one-to-one");
                case PageFieldDisposition.ApplyTaxonomyRelationships:
                    return (IngredientCapability.Available, IngredientDisposition.Transform, "reproduce-reviewed-taxonomy-relationships");
                case PageFieldDisposition.AlreadyHandled:
                    return (IngredientCapability.Available, IngredientDisposition.Preserve, "handled-by-page-writer");
                case PageFieldDisposition.SkipEmpty:
                    return (IngredientCapability.Available, IngredientDisposition.Preserve, "ensure-schema-without-value");
                case PageFieldDisposition.SkipReadOnly:
                case PageFieldDisposition.SkipCalculated:
                case PageFieldDisposition.TargetRuntime:
                    return (IngredientCapability.Available, IngredientDisposition.Substitute, "target-runtime-value");
                case PageFieldDisposition.EvidenceOnly:
                    return (IngredientCapability.Unknown, IngredientDisposition.Delegate, "retain-snapshot");
                case PageFieldDisposition.TargetFieldMissing:
                case PageFieldDisposition.TargetTypeMismatch:
                case PageFieldDisposition.RequiresMapping:
                case PageFieldDisposition.CaptureUnavailable:
                case PageFieldDisposition.Block:
                    return hasValue
                        ? (IngredientCapability.Incompatible, IngredientDisposition.Block, "none")
                        : (IngredientCapability.Unknown, IngredientDisposition.Drop, "discard-no-source-value");
                default:
                    return hasValue
                        ? (IngredientCapability.Unknown, IngredientDisposition.Block, "none")
                        : (IngredientCapability.Unknown, IngredientDisposition.Drop, "discard-no-source-value");
            }
        }

        private static void AddTaxonomyRelationships(
            PublishingPageCaptureBundle snapshot,
            PublishingPageMigrationPlan plan,
            IDictionary<string, PageIngredientAction> actions)
        {
            var planned = (plan.TaxonomyRelationshipActions ?? Array.Empty<TaxonomyRelationshipAction>())
                .ToDictionary(
                    action => RelationshipKey(action.SourceFieldId, action.SourceTermId, action.SourceWssId),
                    StringComparer.Ordinal);
            foreach (var field in snapshot.Fields.Where(field => field != null))
            {
                foreach (var value in (field.TaxonomyValues ?? Array.Empty<PageTaxonomyValueSnapshot>()).Where(value => value != null))
                {
                    Guid termId;
                    Guid.TryParse(value.TermGuid, out termId);
                    var ingredientId = PublishingPageIngredientIds.TaxonomyRelationship(field.Id, termId, value.WssId);
                    planned.TryGetValue(RelationshipKey(field.Id, termId, value.WssId), out var relationshipAction);
                    var executable = relationshipAction != null && relationshipAction.IsExecutable;
                    var evidenceOnly = relationshipAction?.Disposition == TaxonomyRelationshipDisposition.RetainEvidenceOnly;
                    var realization = relationshipAction == null
                        ? "none"
                        : Realization(relationshipAction.Disposition);
                    PublishingPageIngredientActionFactory.Add(actions, PublishingPageIngredientActionFactory.Create(
                        ingredientId,
                        executable
                            ? IngredientCapability.Available
                            : evidenceOnly
                                ? IngredientCapability.Unknown
                                : IngredientCapability.Incompatible,
                        executable
                            ? IngredientDisposition.Transform
                            : evidenceOnly
                                ? IngredientDisposition.Delegate
                                : IngredientDisposition.Block,
                        realization,
                        "policy.taxonomy-relationship.fidelity",
                        relationshipAction?.Reason ?? "No exact target taxonomy relationship action was sealed.",
                        relationshipAction == null || !relationshipAction.IsExecutable
                            ? null
                            : relationshipAction.TargetTermStoreId.ToString("D") + "/" + relationshipAction.TargetBoundTermSetId.ToString("D") + "/" + relationshipAction.SourceTermId.ToString("D"),
                        relationshipAction?.VerificationAssertions?.ToArray() ?? Array.Empty<string>()));
                }
            }
        }

        private static string RelationshipKey(Guid fieldId, Guid termId, int wssId)
        {
            return fieldId.ToString("D") + "/" + termId.ToString("D") + "/" + wssId;
        }

        private static string Realization(TaxonomyRelationshipDisposition disposition)
        {
            switch (disposition)
            {
                case TaxonomyRelationshipDisposition.ReuseLiveInBoundTermSet:
                    return "reuse-live-term-in-mapped-bound-set";
                case TaxonomyRelationshipDisposition.PreserveLiveOutsideBoundTermSet:
                    return "reproduce-live-term-outside-mapped-bound-set";
                case TaxonomyRelationshipDisposition.PreserveDanglingTermAbsent:
                    return "reproduce-dangling-term-with-target-local-wssid";
                case TaxonomyRelationshipDisposition.RetainEvidenceOnly:
                    return "retain-sealed-relationship-evidence";
                default:
                    return "none";
            }
        }
    }
}
