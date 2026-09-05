using System;
using System.Collections.Generic;
using System.Linq;

namespace PnP.Framework.Migration.Taxonomy.Assets.Verification
{
    /// <summary>
    /// Compares the raw reuse, source-Term, membership, and pin relationships
    /// captured for one source Term with a target readback. The owning source
    /// TermSet identity is translated through the reviewed target mapping; every
    /// other captured relationship identity remains exact.
    /// </summary>
    internal static class TaxonomyTermRelationshipFidelity
    {
        public static bool HasCapturedEvidence(TaxonomyTermMaterializationPlan plan)
        {
            return plan != null
                && (plan.SourceIsReused.HasValue
                    || plan.SourceIsSourceTerm.HasValue
                    || plan.SourceReuseSourceTermId.HasValue
                    || (plan.SourceTermSetIds?.Count ?? 0) > 0
                    || plan.SourcePinSourceTermSetId.HasValue);
        }

        public static bool Matches(
            TaxonomyTermMaterializationPlan plan,
            TaxonomyTermTargetProbe observed,
            Guid targetTermSetId,
            out string diagnostic)
        {
            if (plan == null)
            {
                throw new ArgumentNullException(nameof(plan));
            }
            if (observed == null)
            {
                throw new ArgumentNullException(nameof(observed));
            }
            if (targetTermSetId == Guid.Empty)
            {
                throw new ArgumentException("A target TermSet identity is required.", nameof(targetTermSetId));
            }
            if (!HasCapturedEvidence(plan))
            {
                diagnostic = null;
                return true;
            }

            var mismatches = new List<string>();
            if (plan.SourceIsReused != observed.ExistingIsReused)
            {
                mismatches.Add("IsReused expected " + Value(plan.SourceIsReused)
                    + " but observed " + Value(observed.ExistingIsReused));
            }
            if (plan.SourceIsSourceTerm != observed.ExistingIsSourceTerm)
            {
                mismatches.Add("IsSourceTerm expected " + Value(plan.SourceIsSourceTerm)
                    + " but observed " + Value(observed.ExistingIsSourceTerm));
            }
            if (plan.SourceReuseSourceTermId != observed.ExistingReuseSourceTermId)
            {
                mismatches.Add("reported SourceTerm expected " + Value(plan.SourceReuseSourceTermId)
                    + " but observed " + Value(observed.ExistingReuseSourceTermId));
            }

            var capturedMemberships = plan.SourceTermSetIds ?? new List<Guid>();
            if (capturedMemberships.Count > 0)
            {
                var expectedMemberships = MapTermSetIds(plan, targetTermSetId, capturedMemberships);
                var observedMemberships = (observed.ExistingTermSetIds ?? new List<Guid>())
                    .Distinct()
                    .OrderBy(value => value)
                    .ToArray();
                if (!expectedMemberships.SequenceEqual(observedMemberships))
                {
                    mismatches.Add("TermSet memberships expected " + Value(expectedMemberships)
                        + " but observed " + Value(observedMemberships));
                }
            }

            var expectedPinSourceTermSetId = MapTermSetId(
                plan,
                targetTermSetId,
                plan.SourcePinSourceTermSetId);
            if (expectedPinSourceTermSetId != observed.ExistingPinSourceTermSetId)
            {
                mismatches.Add("pin source TermSet expected " + Value(expectedPinSourceTermSetId)
                    + " but observed " + Value(observed.ExistingPinSourceTermSetId));
            }

            diagnostic = mismatches.Count == 0 ? null : string.Join("; ", mismatches) + ".";
            return mismatches.Count == 0;
        }

        public static IList<string> VerificationAssertions(
            TaxonomyTermMaterializationPlan plan,
            TaxonomyTermTargetProbe observed,
            Guid targetTermSetId)
        {
            if (!HasCapturedEvidence(plan))
            {
                return new List<string>();
            }

            var targetTermId = observed?.ResolvedTargetTermId ?? plan.PreferredTargetTermId;
            var assertions = new List<string>
            {
                "Fresh readback proves target Term " + targetTermId.ToString("D")
                    + " reports IsReused=" + Value(plan.SourceIsReused)
                    + " and IsSourceTerm=" + Value(plan.SourceIsSourceTerm) + ".",
                "Fresh readback proves target Term " + targetTermId.ToString("D")
                    + " reports SourceTerm=" + Value(plan.SourceReuseSourceTermId) + "."
            };

            var memberships = plan.SourceTermSetIds ?? new List<Guid>();
            if (memberships.Count > 0)
            {
                assertions.Add(
                    "Fresh readback proves target Term " + targetTermId.ToString("D")
                    + " has exactly these TermSet memberships: "
                    + Value(MapTermSetIds(plan, targetTermSetId, memberships)) + ".");
            }
            assertions.Add(
                "Fresh readback proves target Term " + targetTermId.ToString("D")
                + " reports pin source TermSet="
                + Value(MapTermSetId(plan, targetTermSetId, plan.SourcePinSourceTermSetId)) + ".");
            return assertions;
        }

        private static Guid[] MapTermSetIds(
            TaxonomyTermMaterializationPlan plan,
            Guid targetTermSetId,
            IEnumerable<Guid> sourceTermSetIds)
        {
            return sourceTermSetIds
                .Select(value => value == plan.Source.TermSetId ? targetTermSetId : value)
                .Distinct()
                .OrderBy(value => value)
                .ToArray();
        }

        private static Guid? MapTermSetId(
            TaxonomyTermMaterializationPlan plan,
            Guid targetTermSetId,
            Guid? sourceTermSetId)
        {
            return sourceTermSetId.HasValue && sourceTermSetId.Value == plan.Source.TermSetId
                ? targetTermSetId
                : sourceTermSetId;
        }

        private static string Value(bool? value)
        {
            return value.HasValue ? value.Value.ToString().ToLowerInvariant() : "unavailable";
        }

        private static string Value(Guid? value)
        {
            return value.HasValue ? value.Value.ToString("D") : "none";
        }

        private static string Value(IEnumerable<Guid> values)
        {
            return "[" + string.Join(", ", values.Select(value => value.ToString("D"))) + "]";
        }
    }
}
