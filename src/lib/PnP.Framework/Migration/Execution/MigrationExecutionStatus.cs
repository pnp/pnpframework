namespace PnP.Framework.Migration.Execution
{
    public enum MigrationExecutionStatus
    {
        NotStarted = 1,
        Running = 2,
        Succeeded = 3,
        FailedUnexpectedly = 4,
        PartiallySucceeded = 5
    }
}
