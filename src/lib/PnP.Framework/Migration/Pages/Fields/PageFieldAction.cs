namespace PnP.Framework.Migration.Pages.Fields
{
    public sealed class PageFieldAction
    {
        public string SourceInternalName { get; set; }

        public string TargetInternalName { get; set; }

        public string TargetTypeAsString { get; set; }

        public PageFieldDisposition Disposition { get; set; }

        public string Reason { get; set; }

        public bool WillApply => Disposition == PageFieldDisposition.Apply
            || Disposition == PageFieldDisposition.ApplyTaxonomyRelationships;
    }
}
