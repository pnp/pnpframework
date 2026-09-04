using PnP.Framework.Migration.Pages.Assessment;
using PnP.Framework.Migration.Pages.Ingredients;
using System;
using System.Collections.Generic;
using System.Linq;

namespace PnP.Framework.Migration.Pages.Publishing.Assessment
{
    internal sealed class PublishingPageAssessmentAccumulator
    {
        private readonly IDictionary<string, PageIngredientNode> nodes;
        private readonly IDictionary<string, IList<string>> requiredDependencies;
        private readonly IDictionary<string, PageIngredientAssessment> assessments =
            new Dictionary<string, PageIngredientAssessment>(StringComparer.Ordinal);

        public PublishingPageAssessmentAccumulator(CanonicalPageIngredientGraph graph)
        {
            if (graph == null)
            {
                throw new ArgumentNullException(nameof(graph));
            }

            nodes = graph.Nodes
                .Where(value => value != null)
                .ToDictionary(value => value.Id, StringComparer.Ordinal);
            requiredDependencies = (graph.Edges ?? Array.Empty<PageIngredientEdge>())
                .Where(value => value != null && value.Requirement == PageIngredientRequirement.Required)
                .GroupBy(value => value.FromIngredientId, StringComparer.Ordinal)
                .ToDictionary(
                    group => group.Key,
                    group => (IList<string>)group.Select(value => value.ToIngredientId)
                        .Distinct(StringComparer.Ordinal)
                        .OrderBy(value => value, StringComparer.Ordinal)
                        .ToList(),
                    StringComparer.Ordinal);
        }

        public void Add(
            string ingredientId,
            PageIngredientAssessmentState state,
            IngredientCapability capability,
            IngredientDisposition proposedDisposition,
            string proposedRealization,
            string policyId,
            string reason,
            string targetIdentity = null,
            string mitigationCode = null,
            params string[] verificationAssertions)
        {
            if (!nodes.TryGetValue(ingredientId ?? string.Empty, out var node) || !node.HasContent)
            {
                return;
            }
            var candidate = new PageIngredientAssessment
            {
                IngredientId = ingredientId,
                Kind = node.Kind,
                State = state,
                Capability = capability,
                ProposedDisposition = proposedDisposition,
                ProposedRealization = proposedRealization,
                PolicyId = policyId,
                Reason = reason,
                TargetIdentity = targetIdentity,
                MitigationCode = mitigationCode,
                RequiredDependencyIngredientIds = requiredDependencies.TryGetValue(ingredientId, out var dependencies)
                    ? dependencies.ToList()
                    : new List<string>(),
                VerificationAssertions = (verificationAssertions ?? Array.Empty<string>())
                    .Where(value => !string.IsNullOrWhiteSpace(value))
                    .Distinct(StringComparer.Ordinal)
                    .ToList()
            };
            if (!assessments.TryGetValue(ingredientId, out var existing))
            {
                assessments.Add(ingredientId, candidate);
                return;
            }

            Merge(existing, candidate);
        }

        public IList<PageIngredientAssessment> Complete()
        {
            foreach (var node in nodes.Values.Where(value => value.HasContent).OrderBy(value => value.Id, StringComparer.Ordinal))
            {
                if (!assessments.ContainsKey(node.Id))
                {
                    Add(
                        node.Id,
                        PageIngredientAssessmentState.KnownGap,
                        IngredientCapability.Unknown,
                        IngredientDisposition.Defer,
                        "none",
                        "policy.assessment.handler-missing",
                        "No source-assessment handler classified this captured ingredient.",
                        mitigationCode: "IngredientAssessmentHandlerMissing");
                }
            }

            return assessments.Values.OrderBy(value => value.IngredientId, StringComparer.Ordinal).ToList();
        }

        private static void Merge(PageIngredientAssessment existing, PageIngredientAssessment candidate)
        {
            existing.RequiredDependencyIngredientIds = existing.RequiredDependencyIngredientIds
                .Concat(candidate.RequiredDependencyIngredientIds)
                .Distinct(StringComparer.Ordinal)
                .OrderBy(value => value, StringComparer.Ordinal)
                .ToList();
            existing.VerificationAssertions = existing.VerificationAssertions
                .Concat(candidate.VerificationAssertions)
                .Distinct(StringComparer.Ordinal)
                .ToList();

            if (Equivalent(existing, candidate))
            {
                return;
            }

            var selected = Severity(candidate.State) > Severity(existing.State) ? candidate : existing;
            if (existing.State != candidate.State)
            {
                CopyDecision(existing, selected);
                return;
            }

            existing.State = PageIngredientAssessmentState.KnownGap;
            existing.Capability = IngredientCapability.Incompatible;
            existing.ProposedDisposition = IngredientDisposition.Defer;
            existing.ProposedRealization = "none";
            existing.PolicyId = "policy.assessment.conflict";
            existing.Reason = "Independent source projections produced conflicting decisions: "
                + existing.Reason + " | " + candidate.Reason;
            existing.TargetIdentity = null;
            existing.MitigationCode = "IngredientAssessmentConflict";
            existing.AuthorizationEvidence = null;
        }

        private static bool Equivalent(PageIngredientAssessment left, PageIngredientAssessment right)
        {
            return left.State == right.State
                && left.Capability == right.Capability
                && left.ProposedDisposition == right.ProposedDisposition
                && string.Equals(left.ProposedRealization, right.ProposedRealization, StringComparison.Ordinal)
                && string.Equals(left.PolicyId, right.PolicyId, StringComparison.Ordinal)
                && string.Equals(left.Reason, right.Reason, StringComparison.Ordinal)
                && string.Equals(left.TargetIdentity, right.TargetIdentity, StringComparison.Ordinal)
                && string.Equals(left.MitigationCode, right.MitigationCode, StringComparison.Ordinal)
                && Equivalent(left.AuthorizationEvidence, right.AuthorizationEvidence);
        }

        private static bool Equivalent(
            PageIngredientAuthorizationEvidence left,
            PageIngredientAuthorizationEvidence right)
        {
            if (ReferenceEquals(left, right))
            {
                return true;
            }
            return left != null
                && right != null
                && string.Equals(left.IngredientId, right.IngredientId, StringComparison.Ordinal)
                && string.Equals(left.Operation, right.Operation, StringComparison.Ordinal)
                && string.Equals(left.RequestUri, right.RequestUri, StringComparison.Ordinal)
                && left.HttpStatusCode == right.HttpStatusCode
                && left.ObservedAtUtc == right.ObservedAtUtc
                && string.Equals(left.EvidenceSource, right.EvidenceSource, StringComparison.Ordinal)
                && string.Equals(left.EvidenceSha256, right.EvidenceSha256, StringComparison.OrdinalIgnoreCase);
        }

        private static int Severity(PageIngredientAssessmentState value)
        {
            return value == PageIngredientAssessmentState.AuthorizationBlocked
                ? 4
                : value == PageIngredientAssessmentState.KnownGap
                    ? 3
                    : value == PageIngredientAssessmentState.TargetInspectionRequired ? 2 : 1;
        }

        private static void CopyDecision(PageIngredientAssessment destination, PageIngredientAssessment source)
        {
            destination.State = source.State;
            destination.Capability = source.Capability;
            destination.ProposedDisposition = source.ProposedDisposition;
            destination.ProposedRealization = source.ProposedRealization;
            destination.PolicyId = source.PolicyId;
            destination.Reason = source.Reason;
            destination.TargetIdentity = source.TargetIdentity;
            destination.MitigationCode = source.MitigationCode;
            destination.AuthorizationEvidence = source.AuthorizationEvidence;
        }
    }
}
