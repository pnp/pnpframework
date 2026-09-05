using PnP.Framework.Migration.Taxonomy;

namespace PnP.Framework.Migration.Pages.Fields
{
    public sealed class PageTaxonomyValueSnapshot
    {
        public string Label { get; set; }

        public string TermGuid { get; set; }

        public int WssId { get; set; }

        public TaxonomyValueRelationshipSnapshot Relationship { get; set; }
    }
}
