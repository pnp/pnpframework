namespace PnP.Framework.Migration.Pages.Publishing.Packaging
{
    public enum PublishingPagePackageState
    {
        Draft = 0,
        ApprovalReady = 1,
        /// <summary>
        /// Legacy ambiguous state. New packages never emit it.
        /// </summary>
        Blocked = 2,
        MitigationPending = 3,
        AuthorizationBlocked = 4,
        Invalid = 5
    }
}
