using System;

namespace PnP.Framework.Migration.Taxonomy
{
    public sealed class TaxonomyFieldRelationshipBindingSnapshot
    {
        public Guid FieldId { get; set; }

        public string FieldInternalName { get; set; }

        public Guid TermStoreId { get; set; }

        public Guid BoundTermSetId { get; set; }

        public Guid TextFieldId { get; set; }

        public bool Open { get; set; }
    }
}
