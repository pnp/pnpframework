using System;

namespace PnP.Framework.Migration.Taxonomy
{
    public sealed class TaxonomyTargetMapping
    {
        public Guid SourceTermStoreId { get; set; }

        public Guid SourceTermSetId { get; set; }

        public Guid TargetTermStoreId { get; set; }

        public Guid TargetTermSetId { get; set; }
    }
}
