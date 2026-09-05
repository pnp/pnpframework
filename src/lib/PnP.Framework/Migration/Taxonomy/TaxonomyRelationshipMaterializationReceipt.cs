using System;

namespace PnP.Framework.Migration.Taxonomy
{
    public sealed class TaxonomyRelationshipMaterializationReceipt
    {
        public Guid SourceFieldId { get; set; }

        public Guid SourceTermId { get; set; }

        public int SourceWssId { get; set; }

        public TaxonomyRelationshipDisposition Disposition { get; set; }

        public int TargetValueWssId { get; set; }

        public int TargetTaxCatchAllWssId { get; set; }

        public bool ChangedTarget { get; set; }

        public bool TargetRelationshipStateVerified { get; set; }

        public bool HiddenListIdentityVerified { get; set; }

        public string Message { get; set; }
    }
}
