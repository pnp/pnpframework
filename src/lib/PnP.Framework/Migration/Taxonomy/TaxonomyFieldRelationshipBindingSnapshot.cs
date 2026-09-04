using System;
using System.Text.Json.Serialization;

namespace PnP.Framework.Migration.Taxonomy
{
    public sealed class TaxonomyFieldRelationshipBindingSnapshot
    {
        public Guid FieldId { get; set; }

        public string FieldInternalName { get; set; }

        public Guid TermStoreId { get; set; }

        public Guid BoundTermSetId { get; set; }

        /// <summary>
        /// Gets or sets the source anchor Term identity. <see cref="Guid.Empty"/>
        /// means that the field is bound to the TermSet root.
        /// </summary>
        [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingDefault)]
        public Guid AnchorTermId { get; set; }

        public Guid TextFieldId { get; set; }

        public bool Open { get; set; }
    }
}
