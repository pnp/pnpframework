using System;
using System.Collections.Generic;

namespace PnP.Framework.Migration.Execution
{
    internal sealed class MigrationExecutionRecorder
    {
        private readonly IMigrationExecutionJournal journal;
        private int sequence;

        public MigrationExecutionRecorder(
            Guid operationId,
            string planDigest,
            IMigrationExecutionJournal journal)
        {
            OperationId = operationId;
            PlanDigest = planDigest;
            this.journal = journal ?? NullMigrationExecutionJournal.Instance;
        }

        public Guid OperationId { get; }

        public string PlanDigest { get; }

        public IList<MigrationMutationReceipt> Steps { get; } = new List<MigrationMutationReceipt>();

        public void RecordState(MigrationExecutionStatus status, string message)
        {
            journal.WriteExecutionState(new MigrationExecutionStateReceipt
            {
                OperationId = OperationId,
                PlanDigest = PlanDigest,
                RecordedAtUtc = DateTimeOffset.UtcNow,
                Status = status,
                Message = message
            });
        }

        public T Execute<T>(
            string actionId,
            string description,
            Func<T> action,
            Func<T, MutationOutcome> outcome,
            Func<T, string> message)
        {
            if (string.IsNullOrWhiteSpace(actionId))
            {
                throw new ArgumentException("An action ID is required.", nameof(actionId));
            }

            if (action == null)
            {
                throw new ArgumentNullException(nameof(action));
            }

            var currentSequence = sequence++;
            journal.WriteIntent(new MigrationMutationIntent
            {
                OperationId = OperationId,
                PlanDigest = PlanDigest,
                ActionId = actionId,
                Sequence = currentSequence,
                WrittenAtUtc = DateTimeOffset.UtcNow,
                Description = description
            });
            try
            {
                var result = action();
                RecordReceipt(
                    actionId,
                    currentSequence,
                    outcome == null ? MutationOutcome.Applied : outcome(result),
                    message == null ? description : message(result));
                return result;
            }
            catch (Exception exception)
            {
                RecordReceipt(
                    actionId,
                    currentSequence,
                    MutationOutcome.Failed,
                    exception.Message);
                throw;
            }
        }

        public void Execute(string actionId, string description, Action action)
        {
            Execute(
                actionId,
                description,
                () =>
                {
                    action();
                    return true;
                },
                value => MutationOutcome.Applied,
                value => description);
        }

        public void RecordAlreadySatisfied(string actionId, string message)
        {
            var currentSequence = sequence++;
            RecordReceipt(actionId, currentSequence, MutationOutcome.AlreadySatisfied, message);
        }

        private void RecordReceipt(
            string actionId,
            int currentSequence,
            MutationOutcome outcome,
            string message)
        {
            var receipt = new MigrationMutationReceipt
            {
                OperationId = OperationId,
                PlanDigest = PlanDigest,
                ActionId = actionId,
                Sequence = currentSequence,
                CompletedAtUtc = DateTimeOffset.UtcNow,
                Outcome = outcome,
                Message = message
            };
            Steps.Add(receipt);
            journal.WriteReceipt(receipt);
        }
    }
}
