namespace PnP.Framework.Test.Framework.ExtensibilityCallOut
{
    /// <summary>
    /// This mock simulates an invalid ExtensibilityHandler.
    /// There are at least two situations that will lead to invalid extensibility providers.
    /// 1. The class does not inherit from one of the required interfaces.
    /// 2. If the extensibility provider is built against a different version of the currently
    ///     executing PnP.Framework assembly (for instance in a PowerShell session)
    /// </summary>
    public class ExtensibilityMockInvalidHandler
    {
    }
}
