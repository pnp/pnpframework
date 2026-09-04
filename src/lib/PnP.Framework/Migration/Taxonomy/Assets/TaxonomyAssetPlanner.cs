using PnP.Framework.Migration.Diagnostics;
using PnP.Framework.Migration.Evidence;
using PnP.Framework.Migration.Packaging;
using System;
using System.Collections.Generic;
using System.Linq;

namespace PnP.Framework.Migration.Taxonomy.Assets
{
    public static class TaxonomyAssetPlanner
    {
        public static TaxonomyAssetReviewPlan Create(
            TaxonomyAssetSourceSnapshot source,
            Guid targetTermStoreId)
        {
            if (source == null)
            {
                throw new ArgumentNullException(nameof(source));
            }
            if (targetTermStoreId == Guid.Empty)
            {
                throw new ArgumentException("A target TermStore identity is required.", nameof(targetTermStoreId));
            }

            var result = new TaxonomyAssetReviewPlan
            {
                SourceSnapshotDigest = source.SnapshotDigest,
                TargetTermStoreId = targetTermStoreId
            };
            foreach (var group in (source.TermSets ?? new List<TaxonomyTermSetSourceSnapshot>())
                         .Where(value => value != null && value.Availability == EvidenceAvailability.Captured)
                         .GroupBy(value => value.SourceTenantId.ToString("D") + "/" + value.SourceTermStoreId.ToString("D"), StringComparer.Ordinal)
                         .Select(value => value.First())
                         .OrderBy(value => value.SourceTenantId)
                         .ThenBy(value => value.SourceTermStoreId))
            {
                result.TermGroups.Add(TaxonomyAssetIdentity.CreateTermGroupPlan(
                    group.SourceTenantId,
                    group.SourceTermStoreId,
                    targetTermStoreId));
            }
            foreach (var termSet in (source.TermSets ?? new List<TaxonomyTermSetSourceSnapshot>())
                         .OrderBy(value => value.SourceTermStoreId)
                         .ThenBy(value => value.SourceTermSetId))
            {
                if (termSet.Availability != EvidenceAvailability.Captured)
                {
                    if (termSet.AuthorizationEvidence != null)
                    {
                        LiteralHttpAuthorizationEvidence.Validate(termSet.AuthorizationEvidence);
                        result.AuthorizationStops.Add(termSet.AuthorizationEvidence);
                    }
                    result.Issues.Add(Issue(
                        "SourceTaxonomyTermSetEvidenceUnavailable",
                        "termset:" + termSet.SourceTermStoreId.ToString("D") + "/" + termSet.SourceTermSetId.ToString("D"),
                        string.Join(" ", termSet.Diagnostics ?? new List<string>())));
                    continue;
                }
                result.TermSets.Add(TaxonomyAssetIdentity.CreateTermSetPlan(termSet, targetTermStoreId));
            }

            var plannedSetIds = new HashSet<string>(result.TermSets.Select(Key), StringComparer.Ordinal);
            foreach (var term in OrderTerms(source.Terms))
            {
                if (term.Availability != EvidenceAvailability.Captured)
                {
                    if (term.AuthorizationEvidence != null)
                    {
                        LiteralHttpAuthorizationEvidence.Validate(term.AuthorizationEvidence);
                        result.AuthorizationStops.Add(term.AuthorizationEvidence);
                    }
                    result.Issues.Add(Issue(
                        "SourceTaxonomyTermEvidenceUnavailable",
                        "term:" + term.SourceTermStoreId.ToString("D") + "/" + term.SourceTermSetId.ToString("D") + "/" + term.SourceTermId.ToString("D"),
                        string.Join(" ", term.Diagnostics ?? new List<string>())));
                    continue;
                }
                if (!plannedSetIds.Contains(Key(term.SourceTermStoreId, term.SourceTermSetId)))
                {
                    result.Issues.Add(Issue(
                        "SourceTaxonomyTermSetClosureUnavailable",
                        "term:" + term.SourceTermId.ToString("D"),
                        "The source Term has no captured owning TermSet plan."));
                    continue;
                }
                result.Terms.Add(TaxonomyAssetIdentity.CreateTermPlan(
                    term,
                    targetTermStoreId,
                    term.SourceTermSetId,
                    term.SourceParentTermId));
            }
            result.AuthorizationStops = result.AuthorizationStops
                .GroupBy(value => value.EvidenceSha256, StringComparer.OrdinalIgnoreCase)
                .Select(group => group.First())
                .OrderBy(value => value.RequestUri, StringComparer.OrdinalIgnoreCase)
                .ThenBy(value => value.Operation, StringComparer.Ordinal)
                .ToList();
            result.PlanDigest = ComputeDigest(result);
            return result;
        }

        public static string ComputeDigest(TaxonomyAssetReviewPlan plan)
        {
            if (plan == null)
            {
                throw new ArgumentNullException(nameof(plan));
            }

            return MigrationDigest.ComputeSha256(
                MigrationContractSerializer.SerializeCanonicalWithNullRootProperty(
                    plan,
                    nameof(TaxonomyAssetReviewPlan.PlanDigest)));
        }

        private static IEnumerable<TaxonomyTermSourceSnapshot> OrderTerms(
            IEnumerable<TaxonomyTermSourceSnapshot> terms)
        {
            var remaining = (terms ?? Enumerable.Empty<TaxonomyTermSourceSnapshot>())
                .Where(value => value != null)
                .GroupBy(value => Key(value.SourceTermStoreId, value.SourceTermSetId) + "/" + value.SourceTermId.ToString("D"), StringComparer.Ordinal)
                .Select(group => group.First())
                .ToList();
            var emitted = new HashSet<string>(StringComparer.Ordinal);
            while (remaining.Count > 0)
            {
                var ready = remaining
                    .Where(value => !value.SourceParentTermId.HasValue
                        || emitted.Contains(TermKey(
                            value.SourceTermStoreId,
                            value.SourceTermSetId,
                            value.SourceParentTermId.Value)))
                    .OrderBy(value => value.SourceTermStoreId)
                    .ThenBy(value => value.SourceTermSetId)
                    .ThenBy(value => value.SourceTermId)
                    .ToArray();
                if (ready.Length == 0)
                {
                    foreach (var unresolved in remaining.OrderBy(value => value.SourceTermId))
                    {
                        yield return unresolved;
                    }
                    yield break;
                }
                foreach (var item in ready)
                {
                    remaining.Remove(item);
                    emitted.Add(TermKey(item.SourceTermStoreId, item.SourceTermSetId, item.SourceTermId));
                    yield return item;
                }
            }
        }

        private static string Key(TaxonomyTermSetMaterializationPlan plan)
        {
            return Key(plan.Source.TermStoreId, plan.Source.TermSetId);
        }

        private static string Key(Guid storeId, Guid setId)
        {
            return storeId.ToString("D") + "/" + setId.ToString("D");
        }

        private static string TermKey(Guid storeId, Guid setId, Guid termId)
        {
            return Key(storeId, setId) + "/" + termId.ToString("D");
        }

        private static MigrationIssue Issue(string code, string subject, string message)
        {
            return new MigrationIssue
            {
                Code = code,
                Severity = MigrationIssueSeverity.Error,
                Subject = subject,
                Ingredient = "Taxonomy.Asset",
                Message = string.IsNullOrWhiteSpace(message) ? "Required source taxonomy evidence is unavailable." : message
            };
        }
    }
}
