namespace PnP.Framework.Migration.Taxonomy
{
    public enum TaxonomyRelationshipDisposition
    {
        Block = 0,
        ReuseLiveInBoundTermSet = 1,
        PreserveLiveOutsideBoundTermSet = 2,
        PreserveDanglingTermAbsent = 3,
        RetainEvidenceOnly = 4
    }
}
