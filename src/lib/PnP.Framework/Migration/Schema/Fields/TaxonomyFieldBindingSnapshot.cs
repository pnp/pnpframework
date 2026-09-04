using System;
using System.Text.Json.Serialization;

namespace PnP.Framework.Migration.Schema.Fields
{
    public sealed class TaxonomyFieldBindingSnapshot
    {
        public Guid SourceTermStoreId { get; set; }

        public Guid SourceTermSetId { get; set; }

        /// <summary>
        /// Gets or sets the source anchor Term identity. <see cref="Guid.Empty"/>
        /// means that the field is bound to the TermSet root.
        /// </summary>
        [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingDefault)]
        public Guid AnchorTermId { get; set; }

        public Guid HiddenTextFieldId { get; set; }

        public bool Open { get; set; }
    }
}
