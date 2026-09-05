using System;
using System.Collections.Generic;

namespace PnP.Framework.Migration.Taxonomy
{
    public sealed class TaxonomyRelationshipAction
    {
        public Guid SourceFieldId { get; set; }

        public string SourceFieldInternalName { get; set; }

        public Guid SourceTermId { get; set; }

        public int SourceWssId { get; set; }

        public string SourceEvidenceSha256 { get; set; }

        public TaxonomyRelationshipState SourceState { get; set; }

        public TaxonomyRelationshipDisposition Disposition { get; set; }

        public Guid TargetFieldId { get; set; }

        public Guid TargetTextFieldId { get; set; }

        public bool? TargetFieldOpen { get; set; }

        public Guid TargetTermStoreId { get; set; }

        public Guid TargetBoundTermSetId { get; set; }

        public Guid? TargetLiveTermSetId { get; set; }

        public Guid? TargetValueHiddenListTermSetId { get; set; }

        public Guid? TargetTaxCatchAllHiddenListTermSetId { get; set; }

        public string Reason { get; set; }

        public IList<string> VerificationAssertions { get; set; } = new List<string>();

        public bool IsExecutable => Disposition == TaxonomyRelationshipDisposition.ReuseLiveInBoundTermSet
            || Disposition == TaxonomyRelationshipDisposition.PreserveLiveOutsideBoundTermSet
            || Disposition == TaxonomyRelationshipDisposition.PreserveDanglingTermAbsent;
    }
}
