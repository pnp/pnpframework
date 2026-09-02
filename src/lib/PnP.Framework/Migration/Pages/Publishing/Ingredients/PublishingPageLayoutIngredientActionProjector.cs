using PnP.Framework.Migration.Pages.Ingredients;
using PnP.Framework.Migration.Pages.Publishing.Capture;
using PnP.Framework.Migration.Pages.Publishing.Layouts;
using PnP.Framework.Migration.Pages.Publishing.Planning;
using PnP.Framework.Migration.Schema.Fields;
using System;
using System.Collections.Generic;
using System.Linq;

namespace PnP.Framework.Migration.Pages.Publishing.Ingredients
{
    internal static class PublishingPageLayoutIngredientActionProjector
    {
        public static void Project(
            PublishingPageCaptureBundle snapshot,
            PublishingPageMigrationPlan plan,
            IDictionary<string, PageIngredientAction> actions)
        {
            AddLayoutAndContentType(plan, actions);
            AddContentTypeFields(snapshot, plan, actions);
            AddResources(snapshot, plan, actions);
        }

        private static void AddLayoutAndContentType(
            PublishingPageMigrationPlan plan,
            IDictionary<string, PageIngredientAction> actions)
        {
            var layoutBlocked = plan.LayoutAdmission == null
                || plan.LayoutAdmission.Disposition == PublishingPageLayoutMaterializationDisposition.Block;
            PublishingPageIngredientActionFactory.Add(actions, PublishingPageIngredientActionFactory.Create(
                PublishingPageIngredientIds.Layout,
                layoutBlocked ? IngredientCapability.Incompatible : IngredientCapability.Available,
                layoutBlocked ? IngredientDisposition.Block : IngredientDisposition.Preserve,
                LayoutRealization(plan.LayoutAdmission),
                "policy.layout.publishing",
                plan.LayoutMaterialization?.Reason ?? "The Publishing Page Layout plan is unavailable.",
                plan.LayoutMaterialization?.TargetServerRelativeUrl,
                "The target Page Layout bytes and associated Content Type match the sealed layout plan."));
            PublishingPageIngredientActionFactory.Add(actions, PublishingPageIngredientActionFactory.Create(
                PublishingPageIngredientIds.ContentType,
                layoutBlocked ? IngredientCapability.Incompatible : IngredientCapability.Available,
                layoutBlocked ? IngredientDisposition.Block : IngredientDisposition.Preserve,
                layoutBlocked ? "none" : "materialize-layout-associated-content-type",
                "policy.content-type.layout-association",
                "Use the exact Content Type associated with the approved target Page Layout.",
                plan.TargetProbe?.PageContentTypeId,
                $"The target page ContentTypeId equals '{plan.TargetProbe?.PageContentTypeId}'."));
        }

        private static void AddContentTypeFields(
            PublishingPageCaptureBundle snapshot,
            PublishingPageMigrationPlan plan,
            IDictionary<string, PageIngredientAction> actions)
        {
            var sourceFields = snapshot.Layout?.AssociatedContentTypeSchema?.RequiredFieldClosure
                ?? Array.Empty<FieldSchemaSnapshot>();
            var plannedFields = (plan.LayoutMaterialization?.ContentTypeSchema?.Fields
                    ?? Array.Empty<FieldSchemaMaterializationPlan>())
                .GroupBy(value => value.FieldId)
                .ToDictionary(group => group.Key, group => group.First());
            var reuseStock = plan.LayoutAdmission?.Disposition == PublishingPageLayoutMaterializationDisposition.ReuseTargetStock;
            foreach (var sourceField in sourceFields.Where(value => value != null).GroupBy(value => value.Id).Select(group => group.First()))
            {
                plannedFields.TryGetValue(sourceField.Id, out var fieldPlan);
                var mapping = reuseStock
                    ? (IngredientCapability.Available, IngredientDisposition.Preserve, "reuse-reviewed-stock-schema")
                    : fieldPlan == null
                        ? (IngredientCapability.Incompatible, IngredientDisposition.Block, "none")
                        : Map(fieldPlan.Disposition);
                PublishingPageIngredientActionFactory.Add(actions, PublishingPageIngredientActionFactory.Create(
                    PublishingPageIngredientIds.PageContentTypeField(sourceField.Id),
                    mapping.Item1,
                    mapping.Item2,
                    mapping.Item3,
                    "policy.page-content-type-field." + (reuseStock ? "reviewed-stock" : fieldPlan?.Disposition.ToString().ToLowerInvariant() ?? "missing"),
                    reuseStock
                        ? "The approved stock Page Layout and exact associated Content Type provide this required runtime field schema."
                        : fieldPlan?.Reason ?? "No associated Content Type field materialization decision was produced.",
                    mapping.Item2 == IngredientDisposition.Block
                        ? null
                        : (plan.LayoutMaterialization?.TargetServerRelativeUrl ?? plan.TargetWebUrl) + "#field:" + sourceField.Id.ToString("D"),
                    mapping.Item2 == IngredientDisposition.Block
                        ? null
                        : $"Fresh associated Content Type readback verifies field '{sourceField.InternalName}' under the approved schema policy."));
            }
        }

        private static void AddResources(
            PublishingPageCaptureBundle snapshot,
            PublishingPageMigrationPlan plan,
            IDictionary<string, PageIngredientAction> actions)
        {
            var resourcePlans = (plan.LayoutMaterialization?.ResourceMaterializations
                    ?? Array.Empty<PublishingPageLayoutResourceMaterializationPlan>())
                .Where(value => value != null)
                .GroupBy(value => value.SourceReference ?? string.Empty, StringComparer.Ordinal)
                .ToDictionary(group => group.Key, group => group.First(), StringComparer.Ordinal);
            foreach (var resourceGroup in (snapshot.Layout?.ResourceArtifacts
                         ?? Array.Empty<PublishingPageLayoutResourceSnapshot>())
                     .Where(value => value != null)
                     .GroupBy(value => value.Reference?.Value ?? value.ResolvedSourceUrl ?? string.Empty, StringComparer.Ordinal))
            {
                resourcePlans.TryGetValue(resourceGroup.Key, out var resourcePlan);
                var source = resourceGroup.First();
                var blocked = resourcePlan == null
                    || resourcePlan.Disposition == PublishingPageLayoutResourceMaterializationDisposition.Block;
                var targetRuntime = resourcePlan?.Disposition == PublishingPageLayoutResourceMaterializationDisposition.TargetRuntime;
                PublishingPageIngredientActionFactory.Add(actions, PublishingPageIngredientActionFactory.Create(
                    PublishingPageIngredientIds.LayoutResource(resourceGroup.Key),
                    blocked ? IngredientCapability.Incompatible : IngredientCapability.Available,
                    blocked
                        ? IngredientDisposition.Block
                        : targetRuntime ? IngredientDisposition.Substitute : IngredientDisposition.Preserve,
                    blocked
                        ? "none"
                        : targetRuntime ? "reuse-target-runtime-resource" : "copy-exact-bytes-create-only",
                    "policy.layout.resource",
                    resourcePlan?.Reason ?? "No Page Layout resource materialization decision was produced.",
                    resourcePlan?.TargetReference ?? resourcePlan?.TargetServerRelativeUrl ?? resourcePlan?.SourceReference,
                    blocked
                        ? null
                        : targetRuntime
                            ? "The target-runtime resource reference resolves in the target Publishing runtime."
                            : $"The target resource bytes have SHA-256 '{source.Artifact?.Sha256}'."));
            }
        }

        private static (IngredientCapability Capability, IngredientDisposition Disposition, string Realization) Map(
            FieldSchemaMaterializationDisposition disposition)
        {
            switch (disposition)
            {
                case FieldSchemaMaterializationDisposition.RequireTargetRuntime:
                    return (IngredientCapability.Available, IngredientDisposition.Substitute, "reuse-target-runtime-schema");
                case FieldSchemaMaterializationDisposition.CreateOrReuseOwned:
                    return (IngredientCapability.Available, IngredientDisposition.Preserve, "create-or-reuse-owned-schema");
                default:
                    return (IngredientCapability.Incompatible, IngredientDisposition.Block, "none");
            }
        }

        private static string LayoutRealization(PublishingPageLayoutTargetAdmission admission)
        {
            if (admission == null || admission.Disposition == PublishingPageLayoutMaterializationDisposition.Block)
            {
                return "none";
            }
            return admission.Disposition == PublishingPageLayoutMaterializationDisposition.ReuseTargetStock
                ? "reuse-target"
                : "create-owned";
        }
    }
}
