using System;

namespace PnP.Framework.Migration.Taxonomy
{
    public sealed class TaxonomyRelationshipVerificationResult
    {
        public Guid SourceFieldId { get; set; }

        public string SourceFieldInternalName { get; set; }

        public Guid SourceTermId { get; set; }

        public TaxonomyRelationshipDisposition Disposition { get; set; }

        public int ObservedWssId { get; set; }

        public bool PageValueMatched { get; set; }

        public bool RelationshipStateMatched { get; set; }

        public bool HiddenListIdentityMatched { get; set; }

        public bool TaxCatchAllMatched { get; set; }

        public string Message { get; set; }

        public bool Passed => PageValueMatched
            && RelationshipStateMatched
            && HiddenListIdentityMatched
            && TaxCatchAllMatched;
    }
}
