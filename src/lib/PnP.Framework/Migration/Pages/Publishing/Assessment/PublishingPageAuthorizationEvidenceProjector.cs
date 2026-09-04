using PnP.Framework.Migration.Pages.Assessment;
using PnP.Framework.Migration.Pages.Ingredients;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace PnP.Framework.Migration.Pages.Publishing.Assessment
{
    internal static class PublishingPageAuthorizationEvidenceProjector
    {
        public static void Apply(
            IList<PageIngredientAssessment> assessments,
            PageAssessmentEvidence evidence)
        {
            if (evidence == null)
            {
                return;
            }
            if (!string.Equals(
                    evidence.SchemaVersion,
                    "pnp-page-assessment-evidence/v1",
                    StringComparison.Ordinal))
            {
                throw new InvalidDataException("Unsupported Page assessment evidence schema.");
            }

            var assessmentById = (assessments ?? Array.Empty<PageIngredientAssessment>())
                .Where(value => value != null)
                .ToDictionary(value => value.IngredientId, StringComparer.Ordinal);
            var failures = evidence.AuthorizationFailures
                ?? new List<PageIngredientAuthorizationEvidence>();
            var duplicate = failures
                .Where(value => value != null)
                .GroupBy(value => value.IngredientId, StringComparer.Ordinal)
                .FirstOrDefault(group => string.IsNullOrWhiteSpace(group.Key) || group.Count() != 1);
            if (duplicate != null || failures.Any(value => value == null))
            {
                throw new InvalidDataException(
                    "Ingredient authorization evidence identities are empty or duplicated.");
            }

            foreach (var failure in failures)
            {
                Validate(failure);
                if (!assessmentById.TryGetValue(failure.IngredientId, out var assessment))
                {
                    throw new InvalidDataException(
                        $"Authorization evidence references unknown ingredient '{failure.IngredientId}'.");
                }

                assessment.State = PageIngredientAssessmentState.AuthorizationBlocked;
                assessment.Capability = IngredientCapability.Missing;
                assessment.ProposedDisposition = IngredientDisposition.Block;
                assessment.ProposedRealization = "none";
                assessment.PolicyId = "policy.authorization.literal-http";
                assessment.Reason = $"Source ingredient request '{failure.Operation}' returned literal HTTP {failure.HttpStatusCode}.";
                assessment.TargetIdentity = failure.RequestUri;
                assessment.MitigationCode = null;
                assessment.AuthorizationEvidence = Copy(failure);
                assessment.VerificationAssertions = assessment.VerificationAssertions
                    .Concat(new[]
                    {
                        $"Authorization evidence '{failure.EvidenceSource}' has SHA-256 '{failure.EvidenceSha256}'."
                    })
                    .Distinct(StringComparer.Ordinal)
                    .ToList();
            }
        }

        internal static void Validate(PageIngredientAuthorizationEvidence evidence)
        {
            if (evidence == null
                || string.IsNullOrWhiteSpace(evidence.IngredientId)
                || string.IsNullOrWhiteSpace(evidence.Operation)
                || !Uri.TryCreate(evidence.RequestUri, UriKind.Absolute, out var requestUri)
                || requestUri.Scheme != Uri.UriSchemeHttp && requestUri.Scheme != Uri.UriSchemeHttps
                || evidence.HttpStatusCode != 401 && evidence.HttpStatusCode != 403
                || evidence.ObservedAtUtc == default(DateTimeOffset)
                || string.IsNullOrWhiteSpace(evidence.EvidenceSource)
                || !IsSha256(evidence.EvidenceSha256))
            {
                throw new InvalidDataException(
                    "AuthorizationBlocked requires retained literal HTTP 401/403 wire evidence with operation, URI, timestamp, source, and SHA-256.");
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

        private static PageIngredientAuthorizationEvidence Copy(
            PageIngredientAuthorizationEvidence value)
        {
            return new PageIngredientAuthorizationEvidence
            {
                IngredientId = value.IngredientId,
                Operation = value.Operation,
                RequestUri = value.RequestUri,
                HttpStatusCode = value.HttpStatusCode,
                ObservedAtUtc = value.ObservedAtUtc,
                EvidenceSource = value.EvidenceSource,
                EvidenceSha256 = value.EvidenceSha256
            };
        }
    }
}
