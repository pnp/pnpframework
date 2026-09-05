using PnP.Framework.Migration.Packaging;

namespace PnP.Framework.Migration.Taxonomy.Assets.Packaging
{
    internal static class TaxonomyAssetContractCloner
    {
        public static TaxonomyAssetReviewPlan Clone(TaxonomyAssetReviewPlan value)
        {
            return MigrationContractSerializer.Deserialize<TaxonomyAssetReviewPlan>(
                MigrationContractSerializer.SerializeCanonical(value));
        }
    }
}
