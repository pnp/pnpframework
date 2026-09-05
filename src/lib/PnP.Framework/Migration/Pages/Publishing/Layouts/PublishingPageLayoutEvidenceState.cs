namespace PnP.Framework.Migration.Pages.Publishing.Layouts
{
    public enum PublishingPageLayoutEvidenceState
    {
        Unknown = 0,
        Readable = 1,
        MetadataOnly = 2,
        AccessDenied = 3,
        Missing = 4,
        Failed = 5,
        AuthorizationBlocked = 6
    }
}
