using PnP.Framework.Migration.Pages.Assessment;
using PnP.Framework.Migration.Pages.Ingredients;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace PnP.Framework.Migration.Pages.Publishing.Assessment
{
    public static class PublishingPageMigrationAssessmentValidator
    {
        public static void Validate(PublishingPageMigrationAssessment assessment)
        {
            if (assessment == null)
            {
                throw new ArgumentNullException(nameof(assessment));
            }

            var errors = new List<string>();
            if (!string.Equals(
                    assessment.SchemaVersion,
                    "pnp-publishing-page-migration-assessment/v2",
                    StringComparison.Ordinal))
            {
                errors.Add("Unsupported assessment schema.");
            }
            if (string.IsNullOrWhiteSpace(assessment.SourceSnapshotDigest)
                || string.IsNullOrWhiteSpace(assessment.WorkflowId)
                || string.IsNullOrWhiteSpace(assessment.SelectionDigest)
                || string.IsNullOrWhiteSpace(assessment.SourceWebUrl)
                || string.IsNullOrWhiteSpace(assessment.SourcePageServerRelativeUrl)
                || string.IsNullOrWhiteSpace(assessment.TopologyPlanDigest)
                || assessment.PlanningPolicy == null
                || assessment.IngredientGraph == null)
            {
                errors.Add("Assessment identity, policy, topology, or graph metadata is incomplete.");
            }

            var nodes = assessment.IngredientGraph?.Nodes ?? Array.Empty<PageIngredientNode>();
            var nodeGroups = nodes.Where(value => value != null)
                .GroupBy(value => value.Id, StringComparer.Ordinal).ToArray();
            if (nodeGroups.Any(group => string.IsNullOrWhiteSpace(group.Key) || group.Count() != 1))
            {
                errors.Add("Ingredient graph node identities are empty or duplicated.");
            }
            var nodeById = nodeGroups.Where(group => !string.IsNullOrWhiteSpace(group.Key))
                .ToDictionary(group => group.Key, group => group.First(), StringComparer.Ordinal);
            foreach (var edge in assessment.IngredientGraph?.Edges ?? Array.Empty<PageIngredientEdge>())
            {
                if (edge == null
                    || !nodeById.ContainsKey(edge.FromIngredientId ?? string.Empty)
                    || !nodeById.ContainsKey(edge.ToIngredientId ?? string.Empty))
                {
                    errors.Add("Ingredient graph contains an edge whose endpoint is absent.");
                    break;
                }
            }

            var decisions = assessment.IngredientAssessments ?? Array.Empty<PageIngredientAssessment>();
            var decisionGroups = decisions.Where(value => value != null)
                .GroupBy(value => value.IngredientId, StringComparer.Ordinal).ToArray();
            if (decisionGroups.Any(group => string.IsNullOrWhiteSpace(group.Key) || group.Count() != 1))
            {
                errors.Add("Ingredient assessment identities are empty or duplicated.");
            }
            var decisionById = decisionGroups.Where(group => !string.IsNullOrWhiteSpace(group.Key))
                .ToDictionary(group => group.Key, group => group.First(), StringComparer.Ordinal);
            var contentNodeIds = nodeById.Values.Where(value => value.HasContent)
                .Select(value => value.Id).OrderBy(value => value, StringComparer.Ordinal).ToArray();
            if (!contentNodeIds.SequenceEqual(
                    decisionById.Keys.OrderBy(value => value, StringComparer.Ordinal),
                    StringComparer.Ordinal))
            {
                errors.Add("Every content-bearing graph node must have exactly one assessment and no extra assessment may exist.");
            }
            foreach (var decision in decisionById.Values)
            {
                if (!nodeById.TryGetValue(decision.IngredientId, out var node)
                    || decision.Kind != node.Kind
                    || string.IsNullOrWhiteSpace(decision.PolicyId)
                    || string.IsNullOrWhiteSpace(decision.Reason)
                    || decision.State == PageIngredientAssessmentState.KnownGap
                        && (string.IsNullOrWhiteSpace(decision.MitigationCode)
                            || decision.ProposedDisposition != IngredientDisposition.Defer)
                    || decision.State == PageIngredientAssessmentState.AuthorizationBlocked
                        && decision.ProposedDisposition != IngredientDisposition.Block
                    || decision.State != PageIngredientAssessmentState.AuthorizationBlocked
                        && decision.ProposedDisposition == IngredientDisposition.Block)
                {
                    errors.Add("An ingredient assessment has inconsistent kind, policy, reason, or mitigation metadata.");
                    break;
                }
                if (decision.State == PageIngredientAssessmentState.AuthorizationBlocked)
                {
                    try
                    {
                        PublishingPageAuthorizationEvidenceProjector.Validate(decision.AuthorizationEvidence);
                        if (!string.Equals(
                                decision.IngredientId,
                                decision.AuthorizationEvidence.IngredientId,
                                StringComparison.Ordinal))
                        {
                            errors.Add("Authorization evidence is bound to a different ingredient.");
                            break;
                        }
                    }
                    catch (InvalidDataException exception)
                    {
                        errors.Add(exception.Message);
                        break;
                    }
                }
                else if (decision.AuthorizationEvidence != null)
                {
                    errors.Add("Only AuthorizationBlocked ingredients may retain authorization evidence.");
                    break;
                }
            }

            var expectedState = decisions.Any(value =>
                    value?.State == PageIngredientAssessmentState.AuthorizationBlocked)
                ? PageMigrationAssessmentState.AuthorizationBlocked
                : decisions.Any(value => value?.State == PageIngredientAssessmentState.KnownGap)
                    || (assessment.KnownGaps?.Count ?? 0) > 0
                        ? PageMigrationAssessmentState.KnownGap
                        : PageMigrationAssessmentState.ReadyForTargetInspection;
            if (assessment.State != expectedState)
            {
                errors.Add("Assessment state does not match its ingredient decisions and known blockers.");
            }
            if (string.IsNullOrWhiteSpace(assessment.AssessmentDigest)
                || !string.Equals(
                    assessment.AssessmentDigest,
                    PublishingPageAssessmentDigest.Compute(assessment),
                    StringComparison.OrdinalIgnoreCase))
            {
                errors.Add("Assessment digest is absent or invalid.");
            }

            if (errors.Count > 0)
            {
                throw new InvalidDataException("Invalid Publishing Page migration assessment: " + string.Join(" ", errors));
            }
        }
    }
}
