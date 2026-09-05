using System.Collections.Generic;

namespace PnP.Framework.Migration.Execution
{
    public sealed class InMemoryMigrationExecutionJournal : IMigrationExecutionJournal
    {
        public IList<MigrationExecutionStateReceipt> ExecutionStates { get; } = new List<MigrationExecutionStateReceipt>();

        public IList<MigrationMutationIntent> Intents { get; } = new List<MigrationMutationIntent>();

        public IList<MigrationMutationReceipt> Receipts { get; } = new List<MigrationMutationReceipt>();

        public void WriteExecutionState(MigrationExecutionStateReceipt state)
        {
            ExecutionStates.Add(state);
        }

        public void WriteIntent(MigrationMutationIntent intent)
        {
            Intents.Add(intent);
        }

        public void WriteReceipt(MigrationMutationReceipt receipt)
        {
            Receipts.Add(receipt);
        }
    }
}
