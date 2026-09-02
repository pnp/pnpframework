namespace PnP.Framework.Migration.Pages.Fields
{
    public enum PageFieldDisposition
    {
        Apply = 0,
        AlreadyHandled = 1,
        SkipEmpty = 2,
        SkipReadOnly = 3,
        SkipCalculated = 4,
        TargetFieldMissing = 5,
        TargetTypeMismatch = 6,
        RequiresMapping = 7,
        EvidenceOnly = 8,
        CaptureUnavailable = 9,
        Block = 10,
        ApplyTaxonomyRelationships = 11,
        TargetRuntime = 12
    }
}
