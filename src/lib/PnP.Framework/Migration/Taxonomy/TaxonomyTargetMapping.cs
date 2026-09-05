using System;
using System.Text.Json.Serialization;

namespace PnP.Framework.Migration.Taxonomy
{
    public enum TaxonomyTargetMappingMode
    {
        ResolvedTargetTermSet = 0,
        PreserveUnresolvedSourceReference = 1
    }

    public sealed class TaxonomyTargetMapping
    {
        public Guid SourceTermStoreId { get; set; }

        public Guid SourceTermSetId { get; set; }

        public Guid TargetTermStoreId { get; set; }

        public Guid TargetTermSetId { get; set; }

        [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingDefault)]
        public TaxonomyTargetMappingMode Mode { get; set; }

        [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingDefault)]
        public bool UnresolvedReferenceTargetVerifiedAbsent { get; set; }

        [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
        public string UnresolvedReferenceEvidenceSha256 { get; set; }
    }
}
