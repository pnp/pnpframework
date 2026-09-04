using PnP.Framework.Migration.Pages.Assessment;
using PnP.Framework.Migration.Pages.Ingredients;
using PnP.Framework.Migration.Pages.Publishing.Ingredients;
using PnP.Framework.Migration.Pages.Publishing.Layouts;
using PnP.Framework.Migration.Schema.ContentTypes;
using PnP.Framework.Migration.Schema.Fields;
using System;
using System.Collections.Generic;
using System.Linq;

namespace PnP.Framework.Migration.Pages.Publishing.Assessment
{
    internal static class PublishingPageLayoutAssessmentProjector
    {
        public static void Project(
            PublishingPageAssessmentContext context,
            PublishingPageAssessmentAccumulator assessments)
        {
            AddLayout(context, assessments);
            AddContentType(context, assessments);
            AddContentTypeFields(context, assessments);
            AddResources(context, assessments);
        }

        private static void AddLayout(
            PublishingPageAssessmentContext context,
            PublishingPageAssessmentAccumulator assessments)
        {
            var plan = context.LayoutPlan;
            var failure = context.LayoutPlanningFailure;
            var blocked = failure != null
                || plan == null
                || string.IsNullOrWhiteSpace(plan.TargetServerRelativeUrl)
                || (plan.Disposition != PublishingPageLayoutMaterializationDisposition.ReuseTargetStock
                    && plan.TargetBytes == null);
            var childIngredientPending = !blocked
                && plan.Disposition == PublishingPageLayoutMaterializationDisposition.Block;
            assessments.Add(
                PublishingPageIngredientIds.Layout,
                blocked
                    ? PageIngredientAssessmentState.KnownGap
                    : PageIngredientAssessmentState.TargetInspectionRequired,
                blocked ? IngredientCapability.Incompatible : IngredientCapability.Available,
                blocked ? IngredientDisposition.Defer : IngredientDisposition.Preserve,
                blocked
                    ? "none"
                    : plan.Disposition == PublishingPageLayoutMaterializationDisposition.ReuseTargetStock
                        ? "reuse-reviewed-stock-layout"
                        : "create-or-reuse-digest-owned-layout",
                "policy.layout.publishing",
                failure
                    ?? (plan == null
                        ? "No source-authoritative Page Layout plan was produced."
                        : childIngredientPending
                            ? "The Page Layout bytes, identity, and exact target path have a source-authoritative action; blocked Content Type, field, or resource ingredients remain independently pending."
                            : plan.Reason),
                blocked ? null : plan.TargetServerRelativeUrl,
                blocked ? "PageLayoutMaterializationUnavailable" : null,
                blocked ? null : "Fresh target inspection verifies layout bytes, ownership, registration, zones, and associated Content Type.");
        }

        private static void AddContentType(
            PublishingPageAssessmentContext context,
            PublishingPageAssessmentAccumulator assessments)
        {
            var plan = context.LayoutPlan;
            if (plan?.Disposition == PublishingPageLayoutMaterializationDisposition.ReuseTargetStock)
            {
                assessments.Add(
                    PublishingPageIngredientIds.ContentType,
                    PageIngredientAssessmentState.TargetInspectionRequired,
                    IngredientCapability.Available,
                    IngredientDisposition.Preserve,
                    "reuse-reviewed-stock-content-type",
                    "policy.content-type.layout-association",
                    "Use the exact Content Type associated with the byte-matched reviewed stock Page Layout.",
                    plan.AssociatedContentTypeId,
                    null,
                    "Fresh target inspection verifies the stock Page Layout's exact associated Content Type.");
                return;
            }

            var schema = plan?.ContentTypeSchema;
            var blocked = context.LayoutPlanningFailure != null
                || schema == null
                || IsContentTypeObjectUnavailable(schema);
            var childFieldPending = !blocked
                && schema.Disposition == ContentTypeMaterializationDisposition.Block;
            assessments.Add(
                PublishingPageIngredientIds.ContentType,
                blocked
                    ? PageIngredientAssessmentState.KnownGap
                    : PageIngredientAssessmentState.TargetInspectionRequired,
                blocked ? IngredientCapability.Incompatible : IngredientCapability.Available,
                blocked ? IngredientDisposition.Defer : IngredientDisposition.Preserve,
                blocked
                    ? "none"
                    : schema.Disposition == ContentTypeMaterializationDisposition.ReuseOwned
                        ? "reuse-exact-target-runtime-content-type"
                        : "create-or-reuse-layout-content-type",
                "policy.content-type.layout-association",
                context.LayoutPlanningFailure
                    ?? (schema == null
                        ? "The layout-associated Content Type schema has no source-authoritative materialization plan."
                        : childFieldPending
                            ? "The layout-associated Content Type identity and metadata have a source-authoritative action; blocked child field ingredients remain independently pending."
                            : schema.Reason),
                blocked ? null : schema.ContentTypeId,
                blocked ? "PageLayoutContentTypeMaterializationUnavailable" : null,
                blocked ? null : "Fresh target inspection verifies Content Type identity, parent lineage, field links, and ownership.");
        }

        private static bool IsContentTypeObjectUnavailable(ContentTypeMaterializationPlan plan)
        {
            return plan.Disposition == ContentTypeMaterializationDisposition.Block
                && !(plan.Fields?.Any(value => value?.Disposition == FieldSchemaMaterializationDisposition.Block) ?? false);
        }

        private static void AddContentTypeFields(
            PublishingPageAssessmentContext context,
            PublishingPageAssessmentAccumulator assessments)
        {
            var sourceFields = context.Snapshot.Layout?.AssociatedContentTypeSchema?.RequiredFieldClosure
                ?? Array.Empty<FieldSchemaSnapshot>();
            var stock = context.LayoutPlan?.Disposition == PublishingPageLayoutMaterializationDisposition.ReuseTargetStock;
            var plannedFields = (context.LayoutPlan?.ContentTypeSchema?.Fields
                    ?? Array.Empty<FieldSchemaMaterializationPlan>())
                .Where(value => value != null)
                .GroupBy(value => value.FieldId)
                .ToDictionary(group => group.Key, group => group.First());
            foreach (var sourceField in sourceFields.Where(value => value != null)
                         .GroupBy(value => value.Id).Select(group => group.First()))
            {
                plannedFields.TryGetValue(sourceField.Id, out var fieldPlan);
                var blocked = context.LayoutPlanningFailure != null
                    || (!stock && (fieldPlan == null || fieldPlan.Disposition == FieldSchemaMaterializationDisposition.Block));
                var targetRuntime = !blocked
                    && !stock
                    && fieldPlan.Disposition == FieldSchemaMaterializationDisposition.RequireTargetRuntime;
                assessments.Add(
                    PublishingPageIngredientIds.PageContentTypeField(sourceField.Id),
                    blocked
                        ? PageIngredientAssessmentState.KnownGap
                        : PageIngredientAssessmentState.TargetInspectionRequired,
                    blocked ? IngredientCapability.Incompatible : IngredientCapability.Available,
                    blocked
                        ? IngredientDisposition.Defer
                        : targetRuntime ? IngredientDisposition.Substitute : IngredientDisposition.Preserve,
                    blocked
                        ? "none"
                        : stock ? "reuse-reviewed-stock-schema"
                        : targetRuntime ? "reuse-target-runtime-schema"
                        : "create-or-reuse-owned-schema",
                    "policy.page-content-type-field",
                    context.LayoutPlanningFailure
                        ?? (stock
                            ? "The reviewed stock Page Layout's associated Content Type owns this required field."
                            : fieldPlan?.Reason ?? "No field-schema materialization decision was produced."),
                    blocked
                        ? null
                        : (context.LayoutPlan?.TargetServerRelativeUrl ?? context.TargetWeb?.TargetWebUrl)
                            + "#field:" + sourceField.Id.ToString("D"),
                    blocked ? "PageLayoutFieldMaterializationUnavailable" : null,
                    blocked ? null : $"Fresh target inspection verifies the approved schema policy for field '{sourceField.InternalName}'.");
            }
        }

        private static void AddResources(
            PublishingPageAssessmentContext context,
            PublishingPageAssessmentAccumulator assessments)
        {
            var stock = context.LayoutPlan?.Disposition == PublishingPageLayoutMaterializationDisposition.ReuseTargetStock;
            var plans = (context.LayoutPlan?.ResourceMaterializations
                    ?? Array.Empty<PublishingPageLayoutResourceMaterializationPlan>())
                .Where(value => value != null)
                .GroupBy(value => value.SourceReference ?? string.Empty, StringComparer.Ordinal)
                .ToDictionary(group => group.Key, group => group.First(), StringComparer.Ordinal);
            foreach (var group in (context.Snapshot.Layout?.ResourceArtifacts
                         ?? Array.Empty<PublishingPageLayoutResourceSnapshot>())
                     .Where(value => value != null)
                     .GroupBy(value => value.Reference?.Value ?? value.ResolvedSourceUrl ?? string.Empty, StringComparer.Ordinal))
            {
                plans.TryGetValue(group.Key, out var plan);
                var blocked = context.LayoutPlanningFailure != null
                    || (!stock && (plan == null
                        || plan.Disposition == PublishingPageLayoutResourceMaterializationDisposition.Block));
                var targetRuntime = !blocked
                    && !stock
                    && plan.Disposition == PublishingPageLayoutResourceMaterializationDisposition.TargetRuntime;
                var preserveExternal = !blocked
                    && !stock
                    && plan.Disposition == PublishingPageLayoutResourceMaterializationDisposition.PreserveExternal;
                var source = group.First();
                assessments.Add(
                    PublishingPageIngredientIds.LayoutResource(group.Key),
                    blocked
                        ? PageIngredientAssessmentState.KnownGap
                        : preserveExternal
                            ? PageIngredientAssessmentState.Determined
                            : PageIngredientAssessmentState.TargetInspectionRequired,
                    blocked ? IngredientCapability.Incompatible : IngredientCapability.Available,
                    blocked
                        ? IngredientDisposition.Defer
                        : targetRuntime ? IngredientDisposition.Substitute : IngredientDisposition.Preserve,
                    blocked
                        ? "none"
                        : stock ? "reuse-reviewed-stock-layout-resource"
                        : targetRuntime ? "reuse-target-runtime-resource"
                        : preserveExternal ? "preserve-exact-source-reference"
                        : "copy-exact-bytes-create-only",
                    "policy.layout.resource",
                    context.LayoutPlanningFailure
                        ?? (stock
                            ? "The reviewed byte-matched stock Page Layout owns this exact embedded resource reference."
                            : plan?.Reason ?? "No rendering-resource materialization decision was produced."),
                    blocked
                        ? null
                        : stock ? group.Key : plan.TargetReference ?? plan.TargetServerRelativeUrl ?? plan.SourceReference,
                    blocked ? "PageLayoutResourceMaterializationUnavailable" : null,
                    blocked
                        ? null
                        : targetRuntime
                            ? "Fresh target runtime verification proves the reference resolves."
                            : preserveExternal
                                ? "Fresh Page Layout readback retains the exact authored external reference; no assertion is made about remote payload availability."
                                : $"Fresh readback verifies resource bytes with SHA-256 '{source.Artifact?.Sha256}'.");
            }
        }
    }
}
