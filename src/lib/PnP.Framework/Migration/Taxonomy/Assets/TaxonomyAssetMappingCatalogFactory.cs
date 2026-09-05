using PnP.Framework.Migration.Packaging;
using PnP.Framework.Migration.Taxonomy;
using PnP.Framework.Migration.Taxonomy.Assets.Execution;
using PnP.Framework.Migration.Taxonomy.Assets.Packaging;
using System;
using System.Linq;

namespace PnP.Framework.Migration.Taxonomy.Assets
{
    public static class TaxonomyAssetMappingCatalogFactory
    {
        public static TaxonomyAssetMappingCatalog Create(
            TaxonomyAssetReviewPlan reviewPlan,
            TaxonomyAssetApprovalManifest approval,
            TaxonomyAssetMaterializationReceipt receipt,
            DateTimeOffset generatedAtUtc)
        {
            TaxonomyAssetMaterializationReceiptValidator.Validate(reviewPlan, approval, receipt, true);
            var approvedSets = approval.Actions
                .Where(value => value.Kind == TaxonomyAssetKind.TermSet
                    && value.Decision == TaxonomyAssetApprovalDecision.Approve)
                .ToDictionary(value => value.ActionId, StringComparer.Ordinal);
            var verifiedReceipts = receipt.Actions
                .Where(value => value.Kind == TaxonomyAssetKind.TermSet
                    && value.FreshReadbackPassed)
                .ToDictionary(value => value.ActionId, StringComparer.Ordinal);
            var catalog = new TaxonomyAssetMappingCatalog
            {
                ReviewPlanDigest = reviewPlan.PlanDigest,
                ApprovalDigest = approval.ApprovalDigest,
                MaterializationReceiptDigest = receipt.ReceiptDigest,
                MaterializationOperationId = receipt.OperationId,
                TargetTermStoreId = reviewPlan.TargetTermStoreId,
                GeneratedAtUtc = generatedAtUtc.ToUniversalTime(),
                FieldBindings = approvedSets.Values
                    .OrderBy(value => value.SourceTermStoreId)
                    .ThenBy(value => value.SourceTermSetId)
                    .Select(value =>
                    {
                        var verified = verifiedReceipts[value.ActionId];
                        return new TaxonomyTargetMapping
                        {
                            SourceTermStoreId = value.SourceTermStoreId,
                            SourceTermSetId = value.SourceTermSetId,
                            TargetTermStoreId = verified.TargetTermStoreId,
                            TargetTermSetId = verified.TargetTermSetId,
                            Mode = TaxonomyTargetMappingMode.ResolvedTargetTermSet
                        };
                    })
                    .ToList()
            };
            catalog.CatalogDigest = ComputeDigest(catalog);
            TaxonomyAssetMappingCatalogValidator.Validate(catalog, true);
            return catalog;
        }

        public static string ComputeDigest(TaxonomyAssetMappingCatalog catalog)
        {
            if (catalog == null)
            {
                throw new ArgumentNullException(nameof(catalog));
            }

            return MigrationDigest.ComputeSha256(
                MigrationContractSerializer.SerializeCanonicalWithNullRootProperty(
                    catalog,
                    nameof(TaxonomyAssetMappingCatalog.CatalogDigest)));
        }
    }
}
