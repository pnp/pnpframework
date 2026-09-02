namespace PnP.Framework.Migration.Pages.Ingredients
{
    public enum PageMigrationOutcome
    {
        Unknown = 0,
        Exact = 1,
        ExecutableWithTransform = 2,
        ExecutableWithLoss = 3,
        Blocked = 4
    }
}
