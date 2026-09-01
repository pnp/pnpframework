namespace PnP.Framework.Migration.Execution
{
    internal sealed class NullMigrationExecutionJournal : IMigrationExecutionJournal
    {
        public static readonly NullMigrationExecutionJournal Instance = new NullMigrationExecutionJournal();

        private NullMigrationExecutionJournal()
        {
        }

        public void WriteExecutionState(MigrationExecutionStateReceipt state)
        {
        }

        public void WriteIntent(MigrationMutationIntent intent)
        {
        }

        public void WriteReceipt(MigrationMutationReceipt receipt)
        {
        }
    }
}
