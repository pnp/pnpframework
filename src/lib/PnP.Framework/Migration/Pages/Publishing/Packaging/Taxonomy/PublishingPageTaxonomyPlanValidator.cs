using PnP.Framework.Migration.Pages.Fields;
using PnP.Framework.Migration.Pages.Publishing.Capture;
using PnP.Framework.Migration.Pages.Publishing.Planning;
using PnP.Framework.Migration.Taxonomy;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace PnP.Framework.Migration.Pages.Publishing.Packaging.Taxonomy
{
    internal static class PublishingPageTaxonomyPlanValidator
    {
        public static void Validate(
            PublishingPageCaptureBundle snapshot,
            PublishingPageMigrationPlan plan)
        {
            var sourceRelationships = snapshot.Fields
                .Where(field => field.Kind == PageFieldValueKind.Taxonomy || field.Kind == PageFieldValueKind.TaxonomyCollection)
                .SelectMany(field => field.TaxonomyValues.Select(value => new
                {
                    Field = field,
                    Value = value,
                    Key = TaxonomyKey(field.Id, ParseTermId(value.TermGuid), value.WssId)
                }))
                .ToArray();
            var actions = plan.TaxonomyRelationshipActions.ToArray();
            var duplicate = actions
                .GroupBy(action => action == null ? "<null>" : TaxonomyKey(action.SourceFieldId, action.SourceTermId, action.SourceWssId), StringComparer.Ordinal)
                .FirstOrDefault(group => group.Key == "<null>" || group.Count() > 1);
            if (duplicate != null || actions.Length != sourceRelationships.Length)
            {
                throw new InvalidDataException("The plan must contain exactly one taxonomy relationship action for every captured taxonomy value.");
            }

            var byKey = actions.ToDictionary(
                action => TaxonomyKey(action.SourceFieldId, action.SourceTermId, action.SourceWssId),
                StringComparer.Ordinal);
            var fieldActions = plan.FieldActions.ToDictionary(
                action => action.SourceInternalName,
                StringComparer.OrdinalIgnoreCase);
            foreach (var source in sourceRelationships)
            {
                if (!byKey.TryGetValue(source.Key, out var action)
                    || action.VerificationAssertions == null
                    || action.SourceState != source.Value.Relationship.State
                    || !string.Equals(action.SourceFieldInternalName, source.Field.InternalName, StringComparison.OrdinalIgnoreCase)
                    || !string.Equals(action.SourceEvidenceSha256, source.Value.Relationship.EvidenceSha256, StringComparison.OrdinalIgnoreCase))
                {
                    throw new InvalidDataException($"Taxonomy relationship action '{source.Key}' is not sealed to its exact source evidence.");
                }
                if (action.IsExecutable
                    && (action.TargetFieldId == Guid.Empty
                        || action.TargetTextFieldId == Guid.Empty
                        || !action.TargetFieldOpen.HasValue
                        || action.TargetTermStoreId == Guid.Empty
                        || action.TargetBoundTermSetId == Guid.Empty))
                {
                    throw new InvalidDataException($"Executable taxonomy relationship action '{source.Key}' has no complete target field binding.");
                }

                var fieldAction = fieldActions[source.Field.InternalName];
                if (action.Disposition == TaxonomyRelationshipDisposition.RetainEvidenceOnly)
                {
                    ValidateEvidenceOnly(source.Key, fieldAction, action);
                }
                else if (fieldAction.Disposition != PageFieldDisposition.ApplyTaxonomyRelationships
                    && fieldAction.Disposition != PageFieldDisposition.Block)
                {
                    throw new InvalidDataException($"Taxonomy relationship action '{source.Key}' performs target admission even though its owning field is not selected for replay.");
                }
                ValidateSemantics(source.Field, source.Value, action, plan.PlanningPolicy.TaxonomySchemaMappings);
            }

            foreach (var field in snapshot.Fields.Where(field =>
                         field.Kind == PageFieldValueKind.Taxonomy
                         || field.Kind == PageFieldValueKind.TaxonomyCollection))
            {
                ValidateFieldAction(field, fieldActions[field.InternalName], sourceRelationships
                    .Where(value => value.Field.Id == field.Id)
                    .Select(value => value.Key)
                    .ToArray(), byKey);
            }
        }

        private static void ValidateEvidenceOnly(
            string sourceKey,
            PageFieldAction fieldAction,
            TaxonomyRelationshipAction action)
        {
            if (fieldAction.Disposition == PageFieldDisposition.ApplyTaxonomyRelationships)
            {
                throw new InvalidDataException($"Taxonomy relationship action '{sourceKey}' is evidence-only while its owning field is marked for replay.");
            }
            if (action.TargetTermStoreId != Guid.Empty
                || action.TargetFieldId != Guid.Empty
                || action.TargetTextFieldId != Guid.Empty
                || action.TargetFieldOpen.HasValue
                || action.TargetBoundTermSetId != Guid.Empty
                || action.TargetLiveTermSetId.HasValue
                || action.TargetValueHiddenListTermSetId.HasValue
                || action.TargetTaxCatchAllHiddenListTermSetId.HasValue
                || action.VerificationAssertions.Count > 0)
            {
                throw new InvalidDataException($"Evidence-only taxonomy relationship action '{sourceKey}' must not imply a target binding, materialization, or verification claim.");
            }
        }

        private static void ValidateFieldAction(
            PageFieldValueSnapshot field,
            PageFieldAction fieldAction,
            IReadOnlyCollection<string> fieldKeys,
            IReadOnlyDictionary<string, TaxonomyRelationshipAction> actions)
        {
            var targetFieldBindings = fieldKeys
                .Select(key => actions[key])
                .Where(action => action.Disposition != TaxonomyRelationshipDisposition.RetainEvidenceOnly)
                .Select(action => new
                {
                    action.TargetFieldId,
                    action.TargetTextFieldId,
                    action.TargetFieldOpen,
                    action.TargetTermStoreId,
                    action.TargetBoundTermSetId
                })
                .Distinct()
                .ToArray();
            if (targetFieldBindings.Length > 1)
            {
                throw new InvalidDataException($"Field '{field.InternalName}' has inconsistent target taxonomy field bindings across its relationship actions.");
            }
            if (fieldAction.Disposition == PageFieldDisposition.ApplyTaxonomyRelationships
                && (fieldKeys.Count == 0 || fieldKeys.Any(key => !actions[key].IsExecutable)))
            {
                throw new InvalidDataException($"Field '{field.InternalName}' is marked for taxonomy replay without complete executable relationship actions.");
            }
            if (fieldAction.Disposition == PageFieldDisposition.Block
                && fieldKeys.Count > 0
                && fieldKeys.All(key => actions[key].IsExecutable))
            {
                throw new InvalidDataException($"Field '{field.InternalName}' is blocked even though every taxonomy relationship action is executable.");
            }
            if (fieldAction.Disposition == PageFieldDisposition.Block
                && fieldKeys.Any(key => actions[key].Disposition == TaxonomyRelationshipDisposition.RetainEvidenceOnly)
                && fieldKeys.Any(key => actions[key].Disposition != TaxonomyRelationshipDisposition.RetainEvidenceOnly))
            {
                throw new InvalidDataException($"Field '{field.InternalName}' mixes evidence-only and target-aware taxonomy relationship actions.");
            }
            if (fieldAction.Disposition != PageFieldDisposition.ApplyTaxonomyRelationships
                && fieldAction.Disposition != PageFieldDisposition.Block
                && fieldKeys.Any(key => actions[key].Disposition != TaxonomyRelationshipDisposition.RetainEvidenceOnly))
            {
                throw new InvalidDataException($"Field '{field.InternalName}' is not selected for replay, but one or more taxonomy relationships are not evidence-only.");
            }
        }

        private static void ValidateSemantics(
            PageFieldValueSnapshot field,
            PageTaxonomyValueSnapshot value,
            TaxonomyRelationshipAction action,
            IEnumerable<TaxonomyTargetMapping> mappings)
        {
            var expectedDisposition = value.Relationship.State switch
            {
                TaxonomyRelationshipState.LiveInBoundTermSet => TaxonomyRelationshipDisposition.ReuseLiveInBoundTermSet,
                TaxonomyRelationshipState.LiveOutsideBoundTermSet => TaxonomyRelationshipDisposition.PreserveLiveOutsideBoundTermSet,
                TaxonomyRelationshipState.DanglingTermAbsent => TaxonomyRelationshipDisposition.PreserveDanglingTermAbsent,
                _ => TaxonomyRelationshipDisposition.Block
            };
            if (action.IsExecutable && action.Disposition != expectedDisposition)
            {
                throw new InvalidDataException($"Taxonomy relationship action for '{field.InternalName}:{value.TermGuid}' would change source state '{value.Relationship.State}' into '{action.Disposition}'.");
            }
            if (!action.IsExecutable)
            {
                return;
            }
            if (action.TargetFieldOpen != field.TaxonomyBinding.Open)
            {
                throw new InvalidDataException($"Taxonomy relationship action for '{field.InternalName}:{value.TermGuid}' changes the taxonomy field's open setting.");
            }
            if (!action.VerificationAssertions.Any(item => !string.IsNullOrWhiteSpace(item)))
            {
                throw new InvalidDataException($"Executable taxonomy relationship action for '{field.InternalName}:{value.TermGuid}' has no fresh-readback assertion.");
            }

            var boundMappings = mappings.Where(mapping =>
                    mapping.SourceTermStoreId == field.TaxonomyBinding.TermStoreId
                    && mapping.SourceTermSetId == field.TaxonomyBinding.BoundTermSetId)
                .ToArray();
            if (boundMappings.Length != 1
                || boundMappings[0].TargetTermStoreId != action.TargetTermStoreId
                || boundMappings[0].TargetTermSetId != action.TargetBoundTermSetId)
            {
                throw new InvalidDataException($"Taxonomy relationship action for '{field.InternalName}:{value.TermGuid}' does not use the exact reviewed bound-TermSet mapping.");
            }

            if (action.Disposition == TaxonomyRelationshipDisposition.PreserveDanglingTermAbsent
                && (action.TargetValueHiddenListTermSetId != action.TargetBoundTermSetId
                    || action.TargetTaxCatchAllHiddenListTermSetId != action.TargetBoundTermSetId
                    || action.TargetLiveTermSetId.HasValue))
            {
                throw new InvalidDataException($"Dangling taxonomy action for '{field.InternalName}:{value.TermGuid}' does not preserve the mapped bound-TermSet identity.");
            }
            if (action.Disposition == TaxonomyRelationshipDisposition.PreserveLiveOutsideBoundTermSet)
            {
                var liveMappings = mappings.Where(mapping =>
                        mapping.SourceTermStoreId == field.TaxonomyBinding.TermStoreId
                        && mapping.SourceTermSetId == value.Relationship.LiveTermSetId)
                    .ToArray();
                if (liveMappings.Length != 1
                    || liveMappings[0].TargetTermStoreId != action.TargetTermStoreId
                    || liveMappings[0].TargetTermSetId != action.TargetLiveTermSetId
                    || action.TargetLiveTermSetId == action.TargetBoundTermSetId
                    || !action.TargetValueHiddenListTermSetId.HasValue
                    || !action.TargetTaxCatchAllHiddenListTermSetId.HasValue)
                {
                    throw new InvalidDataException($"Live-outside-bound taxonomy action for '{field.InternalName}:{value.TermGuid}' does not use the exact reviewed bound/live mappings.");
                }
            }
        }

        private static string TaxonomyKey(Guid fieldId, Guid termId, int wssId)
        {
            return fieldId.ToString("D") + "/" + termId.ToString("D") + "/" + wssId;
        }

        private static Guid ParseTermId(string value)
        {
            Guid result;
            return Guid.TryParse(value, out result) ? result : Guid.Empty;
        }
    }
}
