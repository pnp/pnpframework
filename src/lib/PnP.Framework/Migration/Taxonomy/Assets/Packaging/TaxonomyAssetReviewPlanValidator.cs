using PnP.Framework.Migration.Evidence;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace PnP.Framework.Migration.Taxonomy.Assets.Packaging
{
    /// <summary>
    /// Validates the digest-sealed taxonomy asset graph before approval, target
    /// admission, or mutation. Validation never connects to SharePoint.
    /// </summary>
    public static class TaxonomyAssetReviewPlanValidator
    {
        public static void Validate(
            TaxonomyAssetReviewPlan plan,
            bool requireDigest = true,
            bool requireTargetInspection = true)
        {
            if (plan == null)
            {
                throw new ArgumentNullException(nameof(plan));
            }

            var errors = new List<string>();
            if (!string.Equals(plan.SchemaVersion, "pnp-taxonomy-asset-review-plan/v1", StringComparison.Ordinal))
            {
                errors.Add("Unsupported taxonomy asset review-plan schema.");
            }
            if (!IsSha256(plan.SourceSnapshotDigest))
            {
                errors.Add("The source snapshot digest is absent or invalid.");
            }
            if (plan.TargetTermStoreId == Guid.Empty)
            {
                errors.Add("The target TermStore identity is missing.");
            }

            if ((plan.TermGroups ?? new List<TaxonomyTermGroupMaterializationPlan>()).Any(value => value == null)
                || (plan.TermSets ?? new List<TaxonomyTermSetMaterializationPlan>()).Any(value => value == null)
                || (plan.Terms ?? new List<TaxonomyTermMaterializationPlan>()).Any(value => value == null))
            {
                errors.Add("The taxonomy asset plan contains a null materialization plan.");
            }

            var groups = (plan.TermGroups ?? new List<TaxonomyTermGroupMaterializationPlan>())
                .Where(value => value != null)
                .ToArray();
            var groupKeys = new HashSet<string>(StringComparer.Ordinal);
            foreach (var group in groups)
            {
                ValidateGroup(plan, group, groupKeys, errors);
            }

            var sets = (plan.TermSets ?? new List<TaxonomyTermSetMaterializationPlan>())
                .Where(value => value != null)
                .ToArray();
            var setKeys = new HashSet<string>(StringComparer.Ordinal);
            foreach (var set in sets)
            {
                ValidateSet(plan, set, groupKeys, setKeys, errors);
            }

            var terms = (plan.Terms ?? new List<TaxonomyTermMaterializationPlan>())
                .Where(value => value != null)
                .ToArray();
            var termKeys = new HashSet<string>(StringComparer.Ordinal);
            foreach (var term in terms)
            {
                ValidateTerm(plan, term, setKeys, termKeys, errors);
            }
            ValidateParentClosure(terms, termKeys, errors);

            ValidateGroupProbes(plan, groupKeys, requireTargetInspection, errors);
            ValidateSetProbes(plan, setKeys, requireTargetInspection, errors);
            ValidateTermProbes(plan, termKeys, requireTargetInspection, errors);
            ValidateMappingCandidates(plan, setKeys, errors);
            ValidateAuthorizationEvidence(plan, errors);

            if (requireDigest && (!IsSha256(plan.PlanDigest)
                || !string.Equals(plan.PlanDigest, TaxonomyAssetPlanner.ComputeDigest(plan), StringComparison.OrdinalIgnoreCase)))
            {
                errors.Add("The taxonomy asset review-plan digest is absent or invalid.");
            }
            if (errors.Count > 0)
            {
                throw new InvalidDataException("Invalid taxonomy asset review plan: " + string.Join(" ", errors));
            }
        }

        private static void ValidateGroup(
            TaxonomyAssetReviewPlan review,
            TaxonomyTermGroupMaterializationPlan plan,
            ISet<string> keys,
            ICollection<string> errors)
        {
            if (!string.Equals(plan.SchemaVersion, "pnp-taxonomy-termgroup-plan/v1", StringComparison.Ordinal)
                || plan.Source == null
                || plan.Source.TenantId == Guid.Empty
                || plan.Source.TermStoreId == Guid.Empty
                || plan.TargetTermStoreId != review.TargetTermStoreId
                || plan.PreferredTargetGroupId != TaxonomyAssetIdentity.TargetGroupId(plan.Source.TenantId, plan.Source.TermStoreId)
                || !string.Equals(plan.TargetGroupName, TaxonomyAssetIdentity.TargetGroupName, StringComparison.Ordinal)
                || !IsSha256(plan.PlanDigest)
                || !string.Equals(plan.PlanDigest, TaxonomyAssetIdentity.ComputePlanDigest(plan), StringComparison.OrdinalIgnoreCase))
            {
                errors.Add("A TermGroup materialization plan has invalid identity or digest.");
            }
            if (plan.Source != null
                && !keys.Add(TaxonomyAssetApprovalFactory.GroupKey(plan.Source.TenantId, plan.Source.TermStoreId)))
            {
                errors.Add("Duplicate source TermGroup plan '" + plan.Source.TermStoreId.ToString("D") + "'.");
            }
        }

        private static void ValidateSet(
            TaxonomyAssetReviewPlan review,
            TaxonomyTermSetMaterializationPlan plan,
            ISet<string> groupKeys,
            ISet<string> keys,
            ICollection<string> errors)
        {
            var groupKey = plan.Source == null
                ? string.Empty
                : TaxonomyAssetApprovalFactory.GroupKey(plan.Source.TenantId, plan.Source.TermStoreId);
            if (!string.Equals(plan.SchemaVersion, "pnp-taxonomy-termset-plan/v1", StringComparison.Ordinal)
                || plan.Source == null
                || plan.Source.TenantId == Guid.Empty
                || plan.Source.TermStoreId == Guid.Empty
                || plan.Source.TermSetId == Guid.Empty
                || !groupKeys.Contains(groupKey)
                || plan.TargetTermStoreId != review.TargetTermStoreId
                || plan.TargetGroupId != TaxonomyAssetIdentity.TargetGroupId(plan.Source.TenantId, plan.Source.TermStoreId)
                || !string.Equals(plan.TargetGroupName, TaxonomyAssetIdentity.TargetGroupName, StringComparison.Ordinal)
                || plan.PreferredTargetTermSetId == Guid.Empty
                || string.IsNullOrWhiteSpace(plan.SourceTermSetName)
                || string.IsNullOrWhiteSpace(plan.TargetTermSetName)
                || plan.Language <= 0
                || !string.Equals(plan.OriginalIdentifierPropertyName, TaxonomyAssetIdentity.OriginalIdentifierPropertyName, StringComparison.Ordinal)
                || !string.Equals(plan.OriginalIdentifier, TaxonomyAssetIdentity.TermSet(plan.Source), StringComparison.Ordinal)
                || !IsSha256(plan.SourceEvidenceSha256)
                || !IsSha256(plan.PlanDigest)
                || !string.Equals(plan.PlanDigest, TaxonomyAssetIdentity.ComputePlanDigest(plan), StringComparison.OrdinalIgnoreCase))
            {
                errors.Add("A TermSet materialization plan has invalid identity, provenance, or digest.");
            }
            if (plan.Source != null
                && !keys.Add(TaxonomyAssetApprovalFactory.SetKey(plan.Source.TermStoreId, plan.Source.TermSetId)))
            {
                errors.Add("Duplicate source TermSet plan '" + plan.Source.TermSetId.ToString("D") + "'.");
            }
        }

        private static void ValidateGroupProbes(
            TaxonomyAssetReviewPlan plan,
            ISet<string> groupKeys,
            bool required,
            ICollection<string> errors)
        {
            var raw = plan.TermGroupProbes ?? new List<TaxonomyTermGroupTargetProbe>();
            if (raw.Any(value => value == null))
            {
                errors.Add("Target inspection contains a null TermGroup probe.");
            }
            var probes = raw.Where(value => value != null).ToArray();
            var keys = new HashSet<string>(StringComparer.Ordinal);
            foreach (var probe in probes)
            {
                var key = TaxonomyAssetApprovalFactory.GroupKey(probe.SourceTenantId, probe.SourceTermStoreId);
                if (!keys.Add(key)
                    || !groupKeys.Contains(key)
                    || probe.SourceTenantId == Guid.Empty
                    || probe.SourceTermStoreId == Guid.Empty
                    || probe.TargetTermStoreId != plan.TargetTermStoreId
                    || !Enum.IsDefined(typeof(TaxonomyAssetTargetDisposition), probe.Disposition))
                {
                    errors.Add("A TermGroup target probe has duplicate or inconsistent identity.");
                }
                if (required && probe.Disposition == TaxonomyAssetTargetDisposition.TargetInspectionRequired)
                {
                    errors.Add("A TermGroup target probe is still uninspected.");
                }
                ValidateProbeAuthorization(probe.Disposition, probe.AuthorizationEvidence, errors);
            }
            if (required && !keys.SetEquals(groupKeys))
            {
                errors.Add("Target inspection does not cover every TermGroup plan.");
            }
        }

        private static void ValidateTerm(
            TaxonomyAssetReviewPlan review,
            TaxonomyTermMaterializationPlan plan,
            ISet<string> setKeys,
            ISet<string> termKeys,
            ICollection<string> errors)
        {
            var setKey = plan.Source == null
                ? string.Empty
                : TaxonomyAssetApprovalFactory.SetKey(plan.Source.TermStoreId, plan.Source.TermSetId);
            if (!string.Equals(plan.SchemaVersion, "pnp-taxonomy-term-plan/v2", StringComparison.Ordinal)
                || plan.Source == null
                || plan.Source.TenantId == Guid.Empty
                || plan.Source.TermStoreId == Guid.Empty
                || plan.Source.TermSetId == Guid.Empty
                || plan.Source.TermId == Guid.Empty
                || !setKeys.Contains(setKey)
                || plan.TargetTermStoreId != review.TargetTermStoreId
                || plan.TargetTermSetId == Guid.Empty
                || plan.PreferredTargetTermId == Guid.Empty
                || string.IsNullOrWhiteSpace(plan.Name)
                || plan.Language <= 0
                || !string.Equals(plan.OriginalIdentifierPropertyName, TaxonomyAssetIdentity.OriginalIdentifierPropertyName, StringComparison.Ordinal)
                || !string.Equals(plan.OriginalIdentifier, TaxonomyAssetIdentity.Term(plan.Source), StringComparison.Ordinal)
                || (plan.SourceReuseSourceTermId.HasValue && plan.SourceReuseSourceTermId.Value == Guid.Empty)
                || (plan.SourcePinSourceTermSetId.HasValue && plan.SourcePinSourceTermSetId.Value == Guid.Empty)
                || (plan.SourceTermSetIds ?? new List<Guid>()).Any(value => value == Guid.Empty)
                || (plan.SourceTermSetIds ?? new List<Guid>()).Distinct().Count() != (plan.SourceTermSetIds ?? new List<Guid>()).Count
                || !IsSha256(plan.SourceEvidenceSha256)
                || !IsSha256(plan.PlanDigest)
                || !string.Equals(plan.PlanDigest, TaxonomyAssetIdentity.ComputePlanDigest(plan), StringComparison.OrdinalIgnoreCase))
            {
                errors.Add("A Term materialization plan has invalid identity, closure, provenance, or digest.");
            }
            if (plan.Source != null
                && !termKeys.Add(TaxonomyAssetApprovalFactory.TermKey(
                    plan.Source.TermStoreId,
                    plan.Source.TermSetId,
                    plan.Source.TermId)))
            {
                errors.Add("Duplicate source Term plan '" + plan.Source.TermId.ToString("D") + "'.");
            }
        }

        private static void ValidateParentClosure(
            IEnumerable<TaxonomyTermMaterializationPlan> terms,
            ISet<string> termKeys,
            ICollection<string> errors)
        {
            var byKey = terms.Where(value => value.Source != null).ToDictionary(
                value => TaxonomyAssetApprovalFactory.TermKey(
                    value.Source.TermStoreId,
                    value.Source.TermSetId,
                    value.Source.TermId),
                StringComparer.Ordinal);
            foreach (var term in byKey.Values.Where(value => value.TargetParentTermId.HasValue))
            {
                var parentKey = TaxonomyAssetApprovalFactory.TermKey(
                    term.Source.TermStoreId,
                    term.Source.TermSetId,
                    term.TargetParentTermId.Value);
                if (!termKeys.Contains(parentKey)
                    || !byKey.TryGetValue(parentKey, out var parent)
                    || parent.TargetTermSetId != term.TargetTermSetId)
                {
                    errors.Add("Term '" + term.Source.TermId.ToString("D") + "' has no same-set parent plan.");
                }
            }
        }

        private static void ValidateSetProbes(
            TaxonomyAssetReviewPlan plan,
            ISet<string> setKeys,
            bool required,
            ICollection<string> errors)
        {
            var probes = (plan.TermSetProbes ?? new List<TaxonomyTermSetTargetProbe>()).Where(value => value != null).ToArray();
            var keys = new HashSet<string>(StringComparer.Ordinal);
            foreach (var probe in probes)
            {
                var key = TaxonomyAssetApprovalFactory.SetKey(probe.SourceTermStoreId, probe.SourceTermSetId);
                if (!keys.Add(key)
                    || !setKeys.Contains(key)
                    || probe.TargetTermStoreId != plan.TargetTermStoreId
                    || !Enum.IsDefined(typeof(TaxonomyAssetTargetDisposition), probe.Disposition))
                {
                    errors.Add("A TermSet target probe has duplicate or inconsistent identity.");
                }
                if (required && probe.Disposition == TaxonomyAssetTargetDisposition.TargetInspectionRequired)
                {
                    errors.Add("A TermSet target probe is still uninspected.");
                }
                ValidateProbeAuthorization(probe.Disposition, probe.AuthorizationEvidence, errors);
            }
            if (required && !keys.SetEquals(setKeys))
            {
                errors.Add("Target inspection does not cover every TermSet plan.");
            }
        }

        private static void ValidateTermProbes(
            TaxonomyAssetReviewPlan plan,
            ISet<string> termKeys,
            bool required,
            ICollection<string> errors)
        {
            var probes = (plan.TermProbes ?? new List<TaxonomyTermTargetProbe>()).Where(value => value != null).ToArray();
            var keys = new HashSet<string>(StringComparer.Ordinal);
            foreach (var probe in probes)
            {
                var key = TaxonomyAssetApprovalFactory.TermKey(
                    probe.SourceTermStoreId,
                    probe.SourceTermSetId,
                    probe.SourceTermId);
                if (!keys.Add(key)
                    || !termKeys.Contains(key)
                    || probe.TargetTermStoreId != plan.TargetTermStoreId
                    || !Enum.IsDefined(typeof(TaxonomyAssetTargetDisposition), probe.Disposition))
                {
                    errors.Add("A Term target probe has duplicate or inconsistent identity.");
                }
                if ((probe.ExistingReuseSourceTermId.HasValue && probe.ExistingReuseSourceTermId.Value == Guid.Empty)
                    || (probe.ExistingPinSourceTermSetId.HasValue && probe.ExistingPinSourceTermSetId.Value == Guid.Empty)
                    || (probe.ExistingTermSetIds ?? new List<Guid>()).Any(value => value == Guid.Empty)
                    || (probe.ExistingTermSetIds ?? new List<Guid>()).Distinct().Count() != (probe.ExistingTermSetIds ?? new List<Guid>()).Count)
                {
                    errors.Add("A Term target probe has invalid reuse or TermSet membership evidence.");
                }
                if (required && probe.Disposition == TaxonomyAssetTargetDisposition.TargetInspectionRequired)
                {
                    errors.Add("A Term target probe is still uninspected.");
                }
                ValidateProbeAuthorization(probe.Disposition, probe.AuthorizationEvidence, errors);
            }
            if (required && !keys.SetEquals(termKeys))
            {
                errors.Add("Target inspection does not cover every Term plan.");
            }
        }

        private static void ValidateMappingCandidates(
            TaxonomyAssetReviewPlan plan,
            ISet<string> setKeys,
            ICollection<string> errors)
        {
            var keys = new HashSet<string>(StringComparer.Ordinal);
            foreach (var mapping in (plan.MappingCandidates ?? new List<TaxonomyAssetMappingCandidate>()).Where(value => value != null))
            {
                var key = TaxonomyAssetApprovalFactory.SetKey(mapping.SourceTermStoreId, mapping.SourceTermSetId);
                if (!keys.Add(key)
                    || !setKeys.Contains(key)
                    || mapping.TargetTermStoreId != plan.TargetTermStoreId
                    || mapping.TargetTermSetId == Guid.Empty
                    || !IsSha256(mapping.EvidenceSha256))
                {
                    errors.Add("A taxonomy mapping candidate has duplicate or invalid identity/evidence.");
                }
            }
        }

        private static void ValidateAuthorizationEvidence(
            TaxonomyAssetReviewPlan plan,
            ICollection<string> errors)
        {
            foreach (var evidence in plan.AuthorizationStops ?? new List<LiteralHttpAuthorizationEvidence>())
            {
                try
                {
                    LiteralHttpAuthorizationEvidence.Validate(evidence);
                }
                catch (InvalidDataException)
                {
                    errors.Add("Authorization stops contain evidence other than literal HTTP 401/403.");
                }
            }
        }

        private static void ValidateProbeAuthorization(
            TaxonomyAssetTargetDisposition disposition,
            LiteralHttpAuthorizationEvidence evidence,
            ICollection<string> errors)
        {
            if (disposition == TaxonomyAssetTargetDisposition.AuthorizationBlocked && evidence == null)
            {
                errors.Add("An AuthorizationBlocked taxonomy probe lacks literal HTTP 401/403 evidence.");
                return;
            }
            if (evidence == null)
            {
                return;
            }
            try
            {
                LiteralHttpAuthorizationEvidence.Validate(evidence);
            }
            catch (InvalidDataException)
            {
                errors.Add("A taxonomy target probe contains evidence other than literal HTTP 401/403.");
            }
        }

        private static bool IsSha256(string value)
        {
            return !string.IsNullOrWhiteSpace(value)
                && value.Length == 64
                && value.All(character => character >= '0' && character <= '9'
                    || character >= 'a' && character <= 'f'
                    || character >= 'A' && character <= 'F');
        }
    }
}
