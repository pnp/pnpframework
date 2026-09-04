using PnP.Framework.Migration.Pages.Ingredients;
using PnP.Framework.Migration.Pages.Publishing.Capture;
using PnP.Framework.Migration.Pages.Publishing.Layouts;
using PnP.Framework.Migration.Pages.Publishing.Planning;
using PnP.Framework.Migration.Schema.ContentTypes;
using PnP.Framework.Migration.Schema.Fields;
using PnP.Framework.Migration.Taxonomy;
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

            var reuseStock = plan.LayoutAdmission?.Disposition == PublishingPageLayoutMaterializationDisposition.ReuseTargetStock;
            var schemaPlan = plan.LayoutMaterialization?.ContentTypeSchema;
            var schemaAdmission = plan.LayoutAdmission?.ContentTypeSchema;
            var contentTypeBlocked = !reuseStock
                && (schemaPlan == null
                    || schemaPlan.Disposition == ContentTypeMaterializationDisposition.Block
                    || schemaAdmission == null
                    || !schemaAdmission.IsEligible);
            var requireTargetRuntime = !contentTypeBlocked
                && !reuseStock
                && schemaPlan.Disposition == ContentTypeMaterializationDisposition.ReuseOwned;
            PublishingPageIngredientActionFactory.Add(actions, PublishingPageIngredientActionFactory.Create(
                PublishingPageIngredientIds.ContentType,
                contentTypeBlocked ? IngredientCapability.Incompatible : IngredientCapability.Available,
                contentTypeBlocked
                    ? IngredientDisposition.Block
                    : requireTargetRuntime ? IngredientDisposition.Substitute : IngredientDisposition.Preserve,
                contentTypeBlocked
                    ? "none"
                    : reuseStock ? "reuse-reviewed-stock-content-type"
                    : requireTargetRuntime ? "reuse-exact-target-runtime-content-type"
                    : "materialize-layout-associated-content-type",
                "policy.content-type.layout-association",
                reuseStock
                    ? "Use the exact Content Type associated with the approved target stock Page Layout."
                    : schemaPlan?.Reason ?? "No associated Content Type materialization decision was produced.",
                reuseStock
                    ? plan.LayoutTargetProbe?.ResolvedAssociatedContentTypeId
                    : schemaPlan?.ContentTypeId,
                contentTypeBlocked
                    ? null
                    : requireTargetRuntime
                        ? $"Fresh target readback verifies exact existing ContentTypeId '{schemaPlan.ContentTypeId}', metadata, parent lineage, captured field links, and target-runtime fields without schema writes."
                        : $"Fresh target readback verifies associated ContentTypeId '{schemaPlan?.ContentTypeId ?? plan.LayoutTargetProbe?.ResolvedAssociatedContentTypeId}'."));
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
                        : Map(fieldPlan);
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
            var reuseStock = plan.LayoutAdmission?.IsEligible == true
                && plan.LayoutAdmission.Disposition == PublishingPageLayoutMaterializationDisposition.ReuseTargetStock;
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
                var stockResource = reuseStock && !string.IsNullOrWhiteSpace(resourceGroup.Key);
                var blocked = !stockResource && (resourcePlan == null
                    || resourcePlan.Disposition == PublishingPageLayoutResourceMaterializationDisposition.Block);
                var targetRuntime = stockResource
                    || resourcePlan?.Disposition == PublishingPageLayoutResourceMaterializationDisposition.TargetRuntime;
                var preserveExternal = resourcePlan?.Disposition
                    == PublishingPageLayoutResourceMaterializationDisposition.PreserveExternal;
                PublishingPageIngredientActionFactory.Add(actions, PublishingPageIngredientActionFactory.Create(
                    PublishingPageIngredientIds.LayoutResource(resourceGroup.Key),
                    blocked ? IngredientCapability.Incompatible : IngredientCapability.Available,
                    blocked
                        ? IngredientDisposition.Block
                        : targetRuntime ? IngredientDisposition.Substitute : IngredientDisposition.Preserve,
                    blocked
                        ? "none"
                        : stockResource
                            ? "reuse-reviewed-stock-layout-resource"
                            : targetRuntime ? "reuse-target-runtime-resource"
                            : preserveExternal ? "preserve-exact-source-reference"
                            : "copy-exact-bytes-create-only",
                    stockResource ? "policy.layout.resource.reviewed-stock" : "policy.layout.resource",
                    stockResource
                        ? "The admitted target stock Page Layout has exact captured bytes and owns the same embedded resource reference."
                        : resourcePlan?.Reason ?? "No Page Layout resource materialization decision was produced.",
                    stockResource
                        ? resourceGroup.Key
                        : resourcePlan?.TargetReference ?? resourcePlan?.TargetServerRelativeUrl ?? resourcePlan?.SourceReference,
                    blocked
                        ? null
                        : stockResource
                            ? "Fresh readback verifies the exact target stock Page Layout bytes; target runtime verification confirms that its embedded resource reference resolves."
                            : targetRuntime
                            ? "The target-runtime resource reference resolves in the target Publishing runtime."
                            : preserveExternal
                            ? "The target Page Layout retains the exact authored external reference; no assertion is made about remote payload availability."
                            : $"The target resource bytes have SHA-256 '{source.Artifact?.Sha256}'."));
            }
        }

        private static (IngredientCapability Capability, IngredientDisposition Disposition, string Realization) Map(
            FieldSchemaMaterializationPlan field)
        {
            if (field.TaxonomyMappingMode == TaxonomyTargetMappingMode.PreserveUnresolvedSourceReference)
            {
                return (IngredientCapability.Available, IngredientDisposition.Transform, "create-schema-preserving-unresolved-taxonomy-reference");
            }
            switch (field.Disposition)
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
