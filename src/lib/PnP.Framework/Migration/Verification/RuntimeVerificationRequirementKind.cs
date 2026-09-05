namespace PnP.Framework.Migration.Verification
{
    public enum RuntimeVerificationRequirementKind
    {
        PageReachability = 1,
        ErrorShellAbsence = 2,
        AuthoredDomEquality = 3,
        ResourceInventoryEquality = 4,
        ScriptInventoryEquality = 5,
        InlineEventInventoryEquality = 6,
        ScreenshotCapture = 7,
        VisualEquality = 8
    }
}
