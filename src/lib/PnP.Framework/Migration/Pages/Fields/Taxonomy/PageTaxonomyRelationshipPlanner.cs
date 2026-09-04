using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.Taxonomy;
using PnP.Framework.Migration.Pages.Planning;
using PnP.Framework.Migration.Taxonomy;
using System;
using System.Collections.Generic;
using System.Linq;

namespace PnP.Framework.Migration.Pages.Fields.Taxonomy
{
    internal static class PageTaxonomyRelationshipPlanner
    {
        public static List<TaxonomyRelationshipAction> BuildActions(
            ClientContext context,
            List pages,
            IEnumerable<PageFieldValueSnapshot> fields,
            ISet<string> eligibleFieldNames,
            PagePlanningOptions options,
            ICollection<string> blockers,
            ICollection<string> warnings)
        {
            var actions = new List<TaxonomyRelationshipAction>();
            foreach (var field in (fields ?? Array.Empty<PageFieldValueSnapshot>())
                         .Where(IsTaxonomyField)
                         .OrderBy(value => value.InternalName, StringComparer.Ordinal))
            {
                if (eligibleFieldNames == null || !eligibleFieldNames.Contains(field.InternalName))
                {
                    foreach (var value in field.TaxonomyValues.Where(value => value != null))
                    {
                        var retained = CreateBase(field, value, null);
                        retained.Disposition = TaxonomyRelationshipDisposition.RetainEvidenceOnly;
                        retained.Reason = "The owning page field is not selected for replay; its exact taxonomy relationship evidence remains sealed in the snapshot.";
                        actions.Add(retained);
                    }
                    continue;
                }

                var fieldErrors = PageTaxonomyRelationshipEvidence.ValidateSealedField(field);
                var mapping = ResolveMapping(
                    options?.TaxonomySchemaMappings,
                    field.TaxonomyBinding?.TermStoreId ?? Guid.Empty,
                    field.TaxonomyBinding?.BoundTermSetId ?? Guid.Empty,
                    out var mappingError);
                TaxonomyField targetField = null;
                var targetFieldLoaded = false;
                string targetFieldError = null;
                if (fieldErrors.Count == 0 && mapping != null)
                {
                    try
                    {
                        targetField = context.CastTo<TaxonomyField>(pages.Fields.GetByInternalNameOrTitle(field.InternalName));
                        context.Load(targetField,
                            value => value.Id,
                            value => value.InternalName,
                            value => value.SspId,
                            value => value.TermSetId,
                            value => value.TextField,
                            value => value.Open);
                        context.ExecuteQueryRetry();
                        targetFieldLoaded = true;
                        if (targetField.SspId != mapping.TargetTermStoreId
                            || targetField.TermSetId != mapping.TargetTermSetId)
                        {
                            targetFieldError = $"Target field binding is '{targetField.SspId:D}/{targetField.TermSetId:D}', not reviewed mapping '{mapping.TargetTermStoreId:D}/{mapping.TargetTermSetId:D}'.";
                        }
                        else if (targetField.TextField == Guid.Empty)
                        {
                            targetFieldError = "The target taxonomy field has no companion text-field binding.";
                        }
                        else if (targetField.Open != field.TaxonomyBinding.Open)
                        {
                            targetFieldError = $"Target taxonomy field open setting '{targetField.Open}' does not match sealed source setting '{field.TaxonomyBinding.Open}'.";
                        }
                    }
                    catch (Exception exception)
                    {
                        targetFieldError = "Target taxonomy field binding could not be read: " + exception.Message;
                    }
                }

                foreach (var value in field.TaxonomyValues.Where(value => value != null))
                {
                    var action = CreateBase(field, value, mapping);
                    if (targetFieldLoaded)
                    {
                        action.TargetFieldId = targetField.Id;
                        action.TargetTextFieldId = targetField.TextField;
                        action.TargetFieldOpen = targetField.Open;
                    }
                    actions.Add(action);
                    var errors = fieldErrors
                        .Concat(PageTaxonomyRelationshipEvidence.GetFidelityErrors(field, value))
                        .ToList();
                    if (!string.IsNullOrWhiteSpace(mappingError))
                    {
                        errors.Add(mappingError);
                    }
                    if (!string.IsNullOrWhiteSpace(targetFieldError))
                    {
                        errors.Add(targetFieldError);
                    }
                    if (errors.Count > 0)
                    {
                        Block(action, errors, blockers);
                        continue;
                    }

                    try
                    {
                        PlanRelationship(context, field, value, mapping, options, action);
                        var targetErrors = PageTaxonomyRelationshipTargetInspector.InspectHiddenListReadiness(
                            context,
                            value,
                            action);
                        if (targetErrors.Count > 0)
                        {
                            Block(action, targetErrors, blockers);
                        }
                    }
                    catch (Exception exception)
                    {
                        Block(action, new[] { exception.Message }, blockers);
                    }
                }
            }

            foreach (var action in actions.Where(value => value.IsExecutable))
            {
                warnings?.Add($"Taxonomy relationship '{action.SourceFieldInternalName}:{action.SourceTermId:D}' will be reproduced as '{action.Disposition}'; the migration does not create, substitute, or repair Terms.");
            }
            return actions;
        }

        public static IReadOnlyList<string> ValidateFresh(
            ClientContext context,
            IEnumerable<PageFieldValueSnapshot> fields,
            IEnumerable<TaxonomyRelationshipAction> approvedActions,
            PagePlanningOptions options)
        {
            var blockers = new List<string>();
            var approved = (approvedActions ?? Array.Empty<TaxonomyRelationshipAction>()).ToArray();
            if (approved.All(action => action != null
                && action.Disposition == TaxonomyRelationshipDisposition.RetainEvidenceOnly))
            {
                return blockers;
            }

            var pages = context.Web.GetPagesLibrary();
            if (pages == null)
            {
                return new[] { "Fresh taxonomy admission could not find the target Pages library." };
            }
            var eligibleFieldNames = new HashSet<string>(
                approved
                    .Where(action => action.Disposition != TaxonomyRelationshipDisposition.RetainEvidenceOnly)
                    .Select(action => action.SourceFieldInternalName),
                StringComparer.OrdinalIgnoreCase);
            var current = BuildActions(
                context,
                pages,
                fields,
                eligibleFieldNames,
                options,
                blockers,
                new List<string>());
            if (current.Count != approved.Length)
            {
                blockers.Add("The target taxonomy relationship action count changed after approval.");
                return blockers;
            }
            var approvedByKey = approved.ToDictionary(Key, StringComparer.Ordinal);
            foreach (var observed in current)
            {
                if (!approvedByKey.TryGetValue(Key(observed), out var expected)
                    || observed.Disposition != expected.Disposition
                    || observed.TargetTermStoreId != expected.TargetTermStoreId
                    || observed.TargetFieldId != expected.TargetFieldId
                    || observed.TargetTextFieldId != expected.TargetTextFieldId
                    || observed.TargetFieldOpen != expected.TargetFieldOpen
                    || observed.TargetBoundTermSetId != expected.TargetBoundTermSetId
                    || observed.TargetLiveTermSetId != expected.TargetLiveTermSetId
                    || observed.TargetValueHiddenListTermSetId != expected.TargetValueHiddenListTermSetId
                    || observed.TargetTaxCatchAllHiddenListTermSetId != expected.TargetTaxCatchAllHiddenListTermSetId
                    || !string.Equals(observed.SourceEvidenceSha256, expected.SourceEvidenceSha256, StringComparison.OrdinalIgnoreCase))
                {
                    blockers.Add($"Target taxonomy relationship '{observed.SourceFieldInternalName}:{observed.SourceTermId:D}' changed after approval.");
                }
            }
            return blockers.Distinct(StringComparer.Ordinal).OrderBy(value => value, StringComparer.Ordinal).ToArray();
        }

        private static void PlanRelationship(
            ClientContext context,
            PageFieldValueSnapshot field,
            PageTaxonomyValueSnapshot value,
            TaxonomyTargetMapping boundMapping,
            PagePlanningOptions options,
            TaxonomyRelationshipAction action)
        {
            var source = value.Relationship;
            var termId = action.SourceTermId;
            var session = TaxonomySession.GetTaxonomySession(context);
            var store = session.TermStores.GetById(boundMapping.TargetTermStoreId);
            var inBound = store.GetTermInTermSet(boundMapping.TargetTermSetId, termId);
            var global = store.GetTerm(termId);
            context.Load(inBound,
                term => term.Id,
                term => term.Name,
                term => term.PathOfTerm,
                term => term.IsAvailableForTagging);
            context.Load(global,
                term => term.Id,
                term => term.Name,
                term => term.PathOfTerm,
                term => term.IsAvailableForTagging);
            context.ExecuteQueryRetry();
            var inBoundExists = !inBound.ServerObjectIsNull.GetValueOrDefault(true);
            var globalExists = !global.ServerObjectIsNull.GetValueOrDefault(true);

            switch (source.State)
            {
                case TaxonomyRelationshipState.LiveInBoundTermSet:
                    if (!inBoundExists
                        || !string.Equals(inBound.Name, source.LiveTermLabel, StringComparison.Ordinal)
                        || !string.Equals(inBound.PathOfTerm, source.LiveTermPath, StringComparison.Ordinal)
                        || inBound.IsAvailableForTagging != source.LiveTermAvailableForTagging.Value)
                    {
                        throw new InvalidOperationException("The exact live source Term is not present in the reviewed target bound TermSet.");
                    }
                    action.Disposition = TaxonomyRelationshipDisposition.ReuseLiveInBoundTermSet;
                    action.Reason = "Reuse the exact Term GUID in the reviewed target bound TermSet; no Term is created or substituted.";
                    action.VerificationAssertions.Add("Fresh target readback resolves the Term GUID inside the target field's bound TermSet.");
                    action.VerificationAssertions.Add("The persisted page field contains the exact Term GUID.");
                    action.VerificationAssertions.Add("The persisted WssId and TaxCatchAll resolve to the exact target store, bound TermSet, and Term GUID.");
                    break;
                case TaxonomyRelationshipState.DanglingTermAbsent:
                    if (inBoundExists || globalExists)
                    {
                        throw new InvalidOperationException("The source Term is absent, but the target resolves the same GUID live; replay would heal or alter the dangling relationship.");
                    }
                    action.Disposition = TaxonomyRelationshipDisposition.PreserveDanglingTermAbsent;
                    action.TargetValueHiddenListTermSetId = boundMapping.TargetTermSetId;
                    action.TargetTaxCatchAllHiddenListTermSetId = boundMapping.TargetTermSetId;
                    action.Reason = "Recreate the exact dangling relationship with target-local TaxonomyHiddenList identity; keep the Term absent.";
                    action.VerificationAssertions.Add("Fresh target readback proves the Term GUID remains absent from the target term store.");
                    action.VerificationAssertions.Add("The page value and TaxCatchAll resolve through an exact target-local hidden-list row.");
                    break;
                case TaxonomyRelationshipState.LiveOutsideBoundTermSet:
                    if (inBoundExists || !globalExists)
                    {
                        throw new InvalidOperationException("The target does not retain the source live-outside-bound relationship state.");
                    }
                    context.Load(global.TermSet, set => set.Id, set => set.Name);
                    context.ExecuteQueryRetry();
                    var liveMapping = ResolveRequiredLiveMapping(field, source, boundMapping, options);
                    if (global.TermSet.Id != liveMapping.TargetTermSetId
                        || !string.Equals(global.Name, source.LiveTermLabel, StringComparison.Ordinal)
                        || !string.Equals(global.PathOfTerm, source.LiveTermPath, StringComparison.Ordinal)
                        || global.IsAvailableForTagging != source.LiveTermAvailableForTagging.Value)
                    {
                        throw new InvalidOperationException("The target global Term does not exactly match the reviewed source live Term outside the bound TermSet.");
                    }
                    action.Disposition = TaxonomyRelationshipDisposition.PreserveLiveOutsideBoundTermSet;
                    action.TargetLiveTermSetId = liveMapping.TargetTermSetId;
                    action.TargetValueHiddenListTermSetId = MapHiddenListSet(
                        source.ValueHiddenListEntry.TermSetId,
                        field.TaxonomyBinding.BoundTermSetId,
                        source.LiveTermSetId.Value,
                        boundMapping.TargetTermSetId,
                        liveMapping.TargetTermSetId);
                    action.TargetTaxCatchAllHiddenListTermSetId = MapHiddenListSet(
                        source.TaxCatchAllHiddenListEntry.TermSetId,
                        field.TaxonomyBinding.BoundTermSetId,
                        source.LiveTermSetId.Value,
                        boundMapping.TargetTermSetId,
                        liveMapping.TargetTermSetId);
                    action.Reason = "Reproduce the exact live-Term-outside-bound-TermSet relationship and its UNVALIDATED TaxCatchAll row; do not rebind or heal the Term.";
                    action.VerificationAssertions.Add("Fresh target readback proves the Term GUID is live only outside the target field's bound TermSet.");
                    action.VerificationAssertions.Add("TaxCatchAll retains the reviewed UNVALIDATED bound-TermSet hidden-list identity.");
                    break;
                default:
                    throw new InvalidOperationException($"Source taxonomy relationship state '{source.State}' is not executable.");
            }
        }

        private static TaxonomyTargetMapping ResolveRequiredLiveMapping(
            PageFieldValueSnapshot field,
            TaxonomyValueRelationshipSnapshot source,
            TaxonomyTargetMapping boundMapping,
            PagePlanningOptions options)
        {
            var liveMapping = ResolveMapping(
                options.TaxonomySchemaMappings,
                field.TaxonomyBinding.TermStoreId,
                source.LiveTermSetId.Value,
                out var error);
            if (liveMapping == null)
            {
                throw new InvalidOperationException(error);
            }
            if (liveMapping.TargetTermStoreId != boundMapping.TargetTermStoreId)
            {
                throw new InvalidOperationException("The mapped bound and live TermSets must remain in the same target term store.");
            }
            if (liveMapping.TargetTermSetId == boundMapping.TargetTermSetId)
            {
                throw new InvalidOperationException("The mapped bound and live TermSets must remain distinct; merging them would heal the outside-bound relationship.");
            }
            return liveMapping;
        }

        private static Guid MapHiddenListSet(
            Guid sourceSetId,
            Guid sourceBoundSetId,
            Guid sourceLiveSetId,
            Guid targetBoundSetId,
            Guid targetLiveSetId)
        {
            if (sourceSetId == sourceBoundSetId)
            {
                return targetBoundSetId;
            }
            if (sourceSetId == sourceLiveSetId)
            {
                return targetLiveSetId;
            }
            throw new InvalidOperationException("A source hidden-list row belongs to neither the source bound nor source live TermSet.");
        }

        private static TaxonomyRelationshipAction CreateBase(
            PageFieldValueSnapshot field,
            PageTaxonomyValueSnapshot value,
            TaxonomyTargetMapping mapping)
        {
            Guid termId;
            Guid.TryParse(value.TermGuid, out termId);
            return new TaxonomyRelationshipAction
            {
                SourceFieldId = field.Id,
                SourceFieldInternalName = field.InternalName,
                SourceTermId = termId,
                SourceWssId = value.WssId,
                SourceEvidenceSha256 = value.Relationship?.EvidenceSha256,
                SourceState = value.Relationship?.State ?? TaxonomyRelationshipState.Unknown,
                Disposition = TaxonomyRelationshipDisposition.Block,
                TargetTermStoreId = mapping?.TargetTermStoreId ?? Guid.Empty,
                TargetBoundTermSetId = mapping?.TargetTermSetId ?? Guid.Empty
            };
        }

        private static TaxonomyTargetMapping ResolveMapping(
            IEnumerable<TaxonomyTargetMapping> mappings,
            Guid sourceStoreId,
            Guid sourceSetId,
            out string error)
        {
            var matches = (mappings ?? Array.Empty<TaxonomyTargetMapping>())
                .Where(value => value != null
                    && value.SourceTermStoreId == sourceStoreId
                    && value.SourceTermSetId == sourceSetId)
                .ToArray();
            if (matches.Length != 1
                || matches[0].TargetTermStoreId == Guid.Empty
                || matches[0].TargetTermSetId == Guid.Empty)
            {
                error = matches.Length > 1
                    ? $"Multiple reviewed taxonomy mappings exist for source binding '{sourceStoreId:D}/{sourceSetId:D}'."
                    : $"No complete reviewed taxonomy mapping exists for source binding '{sourceStoreId:D}/{sourceSetId:D}'.";
                return null;
            }
            if (matches[0].Mode == TaxonomyTargetMappingMode.PreserveUnresolvedSourceReference)
            {
                error = $"The reviewed taxonomy mapping for source binding '{sourceStoreId:D}/{sourceSetId:D}' preserves an unresolved schema reference and cannot be used to replay live page taxonomy values.";
                return null;
            }
            error = null;
            return matches[0];
        }

        private static void Block(
            TaxonomyRelationshipAction action,
            IEnumerable<string> errors,
            ICollection<string> blockers)
        {
            action.Disposition = TaxonomyRelationshipDisposition.Block;
            action.Reason = string.Join(" ", errors.Where(value => !string.IsNullOrWhiteSpace(value)).Distinct(StringComparer.Ordinal));
            blockers?.Add($"Taxonomy relationship '{action.SourceFieldInternalName}:{action.SourceTermId:D}' is blocked: {action.Reason}");
        }

        private static string Key(TaxonomyRelationshipAction action)
        {
            return action.SourceFieldId.ToString("D") + "/" + action.SourceTermId.ToString("D") + "/" + action.SourceWssId;
        }

        private static bool IsTaxonomyField(PageFieldValueSnapshot field)
        {
            return field != null
                && (field.Kind == PageFieldValueKind.Taxonomy
                    || field.Kind == PageFieldValueKind.TaxonomyCollection);
        }

    }
}
