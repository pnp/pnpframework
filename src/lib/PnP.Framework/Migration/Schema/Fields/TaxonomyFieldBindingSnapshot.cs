using System;

namespace PnP.Framework.Migration.Schema.Fields
{
    public sealed class TaxonomyFieldBindingSnapshot
    {
        public Guid SourceTermStoreId { get; set; }

        public Guid SourceTermSetId { get; set; }

        public Guid HiddenTextFieldId { get; set; }

        public bool Open { get; set; }
    }
}
