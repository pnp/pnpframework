using System;
using System.Text.Json.Serialization;
using PnP.Framework.Migration.Taxonomy;

namespace PnP.Framework.Migration.Schema.Fields
{
    public sealed class FieldSchemaMaterializationPlan
    {
        public Guid FieldId { get; set; }

        public string InternalName { get; set; }

        public string Title { get; set; }

        public string TypeAsString { get; set; }

        public string Group { get; set; }

        public bool Required { get; set; }

        public bool Hidden { get; set; }

        public FieldSchemaRole Role { get; set; }

        public FieldOwnership Ownership { get; set; }

        public FieldSchemaMaterializationDisposition Disposition { get; set; }

        public string SourcePortableSchemaSha256 { get; set; }

        public string TargetSchemaXml { get; set; }

        public string TargetPortableSchemaSha256 { get; set; }

        public Guid? SourceTermStoreId { get; set; }

        public Guid? SourceTermSetId { get; set; }

        [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
        public Guid? SourceAnchorTermId { get; set; }

        public Guid? TargetTermStoreId { get; set; }

        public Guid? TargetTermSetId { get; set; }

        [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingDefault)]
        public TaxonomyTargetMappingMode TaxonomyMappingMode { get; set; }

        [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingDefault)]
        public bool UnresolvedReferenceTargetVerifiedAbsent { get; set; }

        [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
        public string UnresolvedReferenceEvidenceSha256 { get; set; }

        [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
        public Guid? TargetAnchorTermId { get; set; }

        public Guid? HiddenTextFieldId { get; set; }

        public string Reason { get; set; }
    }
}
