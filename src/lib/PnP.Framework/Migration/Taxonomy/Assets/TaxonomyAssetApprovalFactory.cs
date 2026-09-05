using PnP.Framework.Migration.Packaging;
using PnP.Framework.Migration.Taxonomy.Assets.Packaging;
using System;
using System.Collections.Generic;
using System.Linq;

namespace PnP.Framework.Migration.Taxonomy.Assets
{
    public static class TaxonomyAssetApprovalFactory
    {
        public static TaxonomyAssetApprovalManifest CreateTemplate(
            TaxonomyAssetReviewPlan reviewPlan,
            DateTimeOffset generatedAtUtc)
        {
            TaxonomyAssetReviewPlanValidator.Validate(reviewPlan, true, true);
            var manifest = new TaxonomyAssetApprovalManifest
            {
                ReviewPlanDigest = reviewPlan.PlanDigest,
                GeneratedAtUtc = generatedAtUtc.ToUniversalTime()
            };

            var groupPlans = reviewPlan.TermGroups.ToDictionary(
                value => GroupKey(value.Source.TenantId, value.Source.TermStoreId),
                StringComparer.Ordinal);
            foreach (var probe in reviewPlan.TermGroupProbes
                         .OrderBy(value => value.SourceTenantId)
                         .ThenBy(value => value.SourceTermStoreId))
            {
                var plan = groupPlans[GroupKey(probe.SourceTenantId, probe.SourceTermStoreId)];
                manifest.Actions.Add(new TaxonomyAssetActionApproval
                {
                    ActionId = TermGroupActionId(probe.SourceTenantId, probe.SourceTermStoreId),
                    Kind = TaxonomyAssetKind.TermGroup,
                    SourceTenantId = probe.SourceTenantId,
                    SourceTermStoreId = probe.SourceTermStoreId,
                    TargetTermStoreId = probe.TargetTermStoreId,
                    TargetTermGroupId = probe.ResolvedTargetGroupId ?? plan.PreferredTargetGroupId,
                    ReviewedDisposition = probe.Disposition,
                    Decision = probe.Disposition == TaxonomyAssetTargetDisposition.ReuseOwned
                        ? TaxonomyAssetApprovalDecision.Approve
                        : TaxonomyAssetApprovalDecision.Pending,
                    RequiresExplicitReview = probe.Disposition != TaxonomyAssetTargetDisposition.ReuseOwned
                });
            }

            var setPlans = reviewPlan.TermSets.ToDictionary(
                value => SetKey(value.Source.TermStoreId, value.Source.TermSetId),
                StringComparer.Ordinal);
            foreach (var probe in reviewPlan.TermSetProbes
                         .OrderBy(value => value.SourceTermStoreId)
                         .ThenBy(value => value.SourceTermSetId))
            {
                var plan = setPlans[SetKey(probe.SourceTermStoreId, probe.SourceTermSetId)];
                manifest.Actions.Add(new TaxonomyAssetActionApproval
                {
                    ActionId = TermSetActionId(probe.SourceTermStoreId, probe.SourceTermSetId),
                    Kind = TaxonomyAssetKind.TermSet,
                    SourceTenantId = plan.Source.TenantId,
                    SourceTermStoreId = probe.SourceTermStoreId,
                    SourceTermSetId = probe.SourceTermSetId,
                    TargetTermStoreId = probe.TargetTermStoreId,
                    TargetTermGroupId = plan.TargetGroupId,
                    TargetTermSetId = probe.ResolvedTargetTermSetId ?? plan.PreferredTargetTermSetId,
                    ReviewedDisposition = probe.Disposition,
                    Decision = probe.Disposition == TaxonomyAssetTargetDisposition.ReuseOwned
                        ? TaxonomyAssetApprovalDecision.Approve
                        : TaxonomyAssetApprovalDecision.Pending,
                    RequiresExplicitReview = probe.Disposition != TaxonomyAssetTargetDisposition.ReuseOwned
                });
            }

            var termPlans = reviewPlan.Terms.ToDictionary(
                value => TermKey(value.Source.TermStoreId, value.Source.TermSetId, value.Source.TermId),
                StringComparer.Ordinal);
            foreach (var probe in reviewPlan.TermProbes
                         .OrderBy(value => value.SourceTermStoreId)
                         .ThenBy(value => value.SourceTermSetId)
                         .ThenBy(value => value.SourceTermId))
            {
                var plan = termPlans[TermKey(probe.SourceTermStoreId, probe.SourceTermSetId, probe.SourceTermId)];
                manifest.Actions.Add(new TaxonomyAssetActionApproval
                {
                    ActionId = TermActionId(probe.SourceTermStoreId, probe.SourceTermSetId, probe.SourceTermId),
                    Kind = TaxonomyAssetKind.Term,
                    SourceTenantId = plan.Source.TenantId,
                    SourceTermStoreId = probe.SourceTermStoreId,
                    SourceTermSetId = probe.SourceTermSetId,
                    SourceTermId = probe.SourceTermId,
                    TargetTermStoreId = probe.TargetTermStoreId,
                    TargetTermSetId = probe.TargetTermSetId,
                    TargetTermId = probe.ResolvedTargetTermId ?? plan.PreferredTargetTermId,
                    ReviewedDisposition = probe.Disposition,
                    Decision = probe.Disposition == TaxonomyAssetTargetDisposition.ReuseOwned
                        ? TaxonomyAssetApprovalDecision.Approve
                        : TaxonomyAssetApprovalDecision.Pending,
                    RequiresExplicitReview = probe.Disposition != TaxonomyAssetTargetDisposition.ReuseOwned
                });
            }
            return manifest;
        }

        public static void Seal(
            TaxonomyAssetReviewPlan reviewPlan,
            TaxonomyAssetApprovalManifest manifest,
            string approvedBy,
            DateTimeOffset approvedAtUtc)
        {
            if (string.IsNullOrWhiteSpace(approvedBy))
            {
                throw new ArgumentException("An approval identity is required.", nameof(approvedBy));
            }
            manifest.ApprovedBy = approvedBy.Trim();
            manifest.ApprovedAtUtc = approvedAtUtc.ToUniversalTime();
            manifest.ApprovalDigest = null;
            TaxonomyAssetApprovalValidator.Validate(reviewPlan, manifest, false, true);
            manifest.ApprovalDigest = ComputeDigest(manifest);
            TaxonomyAssetApprovalValidator.Validate(reviewPlan, manifest, true, true);
        }

        public static string ComputeDigest(TaxonomyAssetApprovalManifest manifest)
        {
            if (manifest == null)
            {
                throw new ArgumentNullException(nameof(manifest));
            }

            return MigrationDigest.ComputeSha256(
                MigrationContractSerializer.SerializeCanonicalWithNullRootProperty(
                    manifest,
                    nameof(TaxonomyAssetApprovalManifest.ApprovalDigest)));
        }

        public static string TermSetActionId(Guid sourceTermStoreId, Guid sourceTermSetId)
        {
            return "taxonomy.termset." + sourceTermStoreId.ToString("N") + "." + sourceTermSetId.ToString("N");
        }

        public static string TermGroupActionId(Guid sourceTenantId, Guid sourceTermStoreId)
        {
            return "taxonomy.termgroup." + sourceTenantId.ToString("N") + "." + sourceTermStoreId.ToString("N");
        }

        public static string TermActionId(Guid sourceTermStoreId, Guid sourceTermSetId, Guid sourceTermId)
        {
            return "taxonomy.term." + sourceTermStoreId.ToString("N") + "." + sourceTermSetId.ToString("N") + "." + sourceTermId.ToString("N");
        }

        internal static string SetKey(Guid storeId, Guid setId)
        {
            return storeId.ToString("D") + "/" + setId.ToString("D");
        }

        internal static string GroupKey(Guid tenantId, Guid storeId)
        {
            return tenantId.ToString("D") + "/" + storeId.ToString("D");
        }

        internal static string TermKey(Guid storeId, Guid setId, Guid termId)
        {
            return SetKey(storeId, setId) + "/" + termId.ToString("D");
        }
    }
}
