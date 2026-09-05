using PnP.Framework.Migration.Packaging;
using PnP.Framework.Migration.Taxonomy;
using PnP.Framework.Migration.Taxonomy.Assets;
using PnP.Framework.Migration.Taxonomy.Assets.Packaging;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace PnP.Framework.Migration.Pages.Planning
{
    /// <summary>
    /// Normalizes verified taxonomy mappings into the existing planning surface
    /// while retaining the complete catalog evidence in the sealed page plan.
    /// </summary>
    public static class PagePlanningTaxonomyMappingResolver
    {
        public static void UseVerifiedCatalog(
            PagePlanningOptions options,
            TaxonomyAssetMappingCatalog catalog)
        {
            if (options == null)
            {
                throw new ArgumentNullException(nameof(options));
            }
            TaxonomyAssetMappingCatalogValidator.Validate(catalog, true);
            options.TaxonomyAssetMappingCatalog = Clone(catalog);
            options.TaxonomySchemaMappings = CloneMappings(catalog.FieldBindings);
        }

        public static void Normalize(PagePlanningOptions options)
        {
            if (options == null)
            {
                throw new ArgumentNullException(nameof(options));
            }
            if (options.TaxonomyAssetMappingCatalog == null)
            {
                return;
            }

            TaxonomyAssetMappingCatalogValidator.Validate(options.TaxonomyAssetMappingCatalog, true);
            var verified = options.TaxonomyAssetMappingCatalog.FieldBindings ?? new List<TaxonomyTargetMapping>();
            var supplied = options.TaxonomySchemaMappings ?? new List<TaxonomyTargetMapping>();
            if (supplied.Count > 0 && !Equivalent(supplied, verified))
            {
                throw new InvalidDataException(
                    "The page planning taxonomy mappings differ from the digest-sealed taxonomy asset mapping catalog.");
            }
            options.TaxonomySchemaMappings = CloneMappings(verified);
        }

        /// <summary>
        /// Produces assessment-only prospective mappings from a validated
        /// taxonomy target preflight. Every source set covered by the preflight
        /// replaces any older supplied guess. Only deterministic owned reuse,
        /// create, or owned-drift reconciliation candidates are exposed; an
        /// external reuse, collision, retry, uninspected, or authorization state
        /// remains in the nonterminal mitigation queue.
        /// </summary>
        public static IList<TaxonomyTargetMapping> ResolveForAssessment(
            IEnumerable<TaxonomyTargetMapping> supplied,
            TaxonomyAssetReviewPlan reviewPlan)
        {
            if (reviewPlan == null)
            {
                return CloneMappings(supplied);
            }

            TaxonomyAssetReviewPlanValidator.Validate(reviewPlan, true, true);
            var candidates = (reviewPlan.MappingCandidates ?? new List<TaxonomyAssetMappingCandidate>())
                .Where(value => value != null)
                .ToArray();
            var reviewedSourceKeys = new HashSet<string>(
                candidates.Select(value => SourceKey(value.SourceTermStoreId, value.SourceTermSetId)),
                StringComparer.Ordinal);
            var result = (supplied ?? Enumerable.Empty<TaxonomyTargetMapping>())
                .Where(value => value != null
                    && !reviewedSourceKeys.Contains(SourceKey(value.SourceTermStoreId, value.SourceTermSetId)))
                .Select(CloneMapping)
                .ToList();

            result.AddRange(candidates
                .Where(value => IsDeterministicAssessmentCandidate(value.Disposition))
                .Select(value => new TaxonomyTargetMapping
                {
                    SourceTermStoreId = value.SourceTermStoreId,
                    SourceTermSetId = value.SourceTermSetId,
                    TargetTermStoreId = value.TargetTermStoreId,
                    TargetTermSetId = value.TargetTermSetId,
                    Mode = TaxonomyTargetMappingMode.ResolvedTargetTermSet
                }));

            var duplicate = result
                .GroupBy(value => SourceKey(value.SourceTermStoreId, value.SourceTermSetId), StringComparer.Ordinal)
                .FirstOrDefault(group => group.Count() != 1);
            if (duplicate != null)
            {
                throw new InvalidDataException(
                    "Assessment taxonomy mappings contain more than one target for source '"
                    + duplicate.Key + "'.");
            }

            return result
                .OrderBy(value => value.SourceTermStoreId)
                .ThenBy(value => value.SourceTermSetId)
                .ToList();
        }

        public static TaxonomyAssetMappingCandidate FindAssessmentCandidate(
            TaxonomyAssetReviewPlan reviewPlan,
            Guid sourceTermStoreId,
            Guid sourceTermSetId)
        {
            if (reviewPlan == null)
            {
                return null;
            }

            return (reviewPlan.MappingCandidates ?? new List<TaxonomyAssetMappingCandidate>())
                .SingleOrDefault(value => value != null
                    && value.SourceTermStoreId == sourceTermStoreId
                    && value.SourceTermSetId == sourceTermSetId);
        }

        private static bool IsDeterministicAssessmentCandidate(
            TaxonomyAssetTargetDisposition disposition)
        {
            return disposition == TaxonomyAssetTargetDisposition.ReuseOwned
                || disposition == TaxonomyAssetTargetDisposition.CreateMissing
                || disposition == TaxonomyAssetTargetDisposition.ReconcileOwnedPlanDrift;
        }

        private static string SourceKey(Guid termStoreId, Guid termSetId)
        {
            return termStoreId.ToString("D") + "/" + termSetId.ToString("D");
        }

        private static TaxonomyTargetMapping CloneMapping(TaxonomyTargetMapping value)
        {
            return new TaxonomyTargetMapping
            {
                SourceTermStoreId = value.SourceTermStoreId,
                SourceTermSetId = value.SourceTermSetId,
                TargetTermStoreId = value.TargetTermStoreId,
                TargetTermSetId = value.TargetTermSetId,
                Mode = value.Mode,
                UnresolvedReferenceTargetVerifiedAbsent = value.UnresolvedReferenceTargetVerifiedAbsent,
                UnresolvedReferenceEvidenceSha256 = value.UnresolvedReferenceEvidenceSha256
            };
        }

        public static TaxonomyAssetMappingCatalog Clone(TaxonomyAssetMappingCatalog catalog)
        {
            if (catalog == null)
            {
                return null;
            }
            return MigrationContractSerializer.Deserialize<TaxonomyAssetMappingCatalog>(
                MigrationContractSerializer.SerializeCanonical(catalog));
        }

        private static bool Equivalent(
            IEnumerable<TaxonomyTargetMapping> left,
            IEnumerable<TaxonomyTargetMapping> right)
        {
            return Canonical(left).SequenceEqual(Canonical(right), StringComparer.Ordinal);
        }

        private static IEnumerable<string> Canonical(IEnumerable<TaxonomyTargetMapping> values)
        {
            return (values ?? Enumerable.Empty<TaxonomyTargetMapping>())
                .Where(value => value != null)
                .Select(value => value.SourceTermStoreId.ToString("D") + "/"
                    + value.SourceTermSetId.ToString("D") + "->"
                    + value.TargetTermStoreId.ToString("D") + "/"
                    + value.TargetTermSetId.ToString("D") + "/"
                    + value.Mode.ToString() + "/"
                    + value.UnresolvedReferenceTargetVerifiedAbsent + "/"
                    + (value.UnresolvedReferenceEvidenceSha256 ?? string.Empty))
                .OrderBy(value => value, StringComparer.Ordinal);
        }

        private static IList<TaxonomyTargetMapping> CloneMappings(IEnumerable<TaxonomyTargetMapping> values)
        {
            return (values ?? Enumerable.Empty<TaxonomyTargetMapping>())
                .Where(value => value != null)
                .Select(CloneMapping)
                .ToList();
        }
    }
}
