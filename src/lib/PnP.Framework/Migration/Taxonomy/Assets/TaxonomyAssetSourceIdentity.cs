using System;

namespace PnP.Framework.Migration.Taxonomy.Assets
{
    public sealed class TaxonomyTermGroupSourceIdentity
    {
        public Guid TenantId { get; set; }

        public Guid TermStoreId { get; set; }
    }

    public sealed class TaxonomyTermSetSourceIdentity
    {
        public Guid TenantId { get; set; }

        public Guid TermStoreId { get; set; }

        public Guid TermSetId { get; set; }
    }

    public sealed class TaxonomyTermSourceIdentity
    {
        public Guid TenantId { get; set; }

        public Guid TermStoreId { get; set; }

        public Guid TermSetId { get; set; }

        public Guid TermId { get; set; }
    }
}
