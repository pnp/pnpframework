using PnP.Framework.Migration.Taxonomy;
using System;
using System.Collections.Generic;

namespace PnP.Framework.Migration.Taxonomy.Assets
{
    /// <summary>
    /// A page/schema-planning mapping catalog derived only from a sealed
    /// taxonomy approval and its successful, digest-sealed fresh-readback receipt.
    /// </summary>
    public sealed class TaxonomyAssetMappingCatalog
    {
        public string SchemaVersion { get; set; } = "pnp-taxonomy-asset-mapping-catalog/v1";

        public string ReviewPlanDigest { get; set; }

        public string ApprovalDigest { get; set; }

        public string MaterializationReceiptDigest { get; set; }

        public Guid MaterializationOperationId { get; set; }

        public Guid TargetTermStoreId { get; set; }

        public DateTimeOffset GeneratedAtUtc { get; set; }

        public IList<TaxonomyTargetMapping> FieldBindings { get; set; } = new List<TaxonomyTargetMapping>();

        public string CatalogDigest { get; set; }
    }
}
