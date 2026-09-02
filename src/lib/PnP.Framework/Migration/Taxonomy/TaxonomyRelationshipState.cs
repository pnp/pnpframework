namespace PnP.Framework.Migration.Taxonomy
{
    public enum TaxonomyRelationshipState
    {
        Unknown = 0,
        LiveInBoundTermSet = 1,
        LiveOutsideBoundTermSet = 2,
        DanglingTermAbsent = 3,
        Conflict = 4
    }
}
