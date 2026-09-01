namespace PnP.Framework.Migration.Execution
{
    public interface IMigrationExecutionJournal
    {
        void WriteExecutionState(MigrationExecutionStateReceipt state);

        void WriteIntent(MigrationMutationIntent intent);

        void WriteReceipt(MigrationMutationReceipt receipt);
    }
}
