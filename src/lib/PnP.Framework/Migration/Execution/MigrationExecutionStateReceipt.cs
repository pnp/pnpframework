using System;

namespace PnP.Framework.Migration.Execution
{
    public sealed class MigrationExecutionStateReceipt
    {
        public Guid OperationId { get; set; }

        public string PlanDigest { get; set; }

        public DateTimeOffset RecordedAtUtc { get; set; }

        public MigrationExecutionStatus Status { get; set; }

        public string Message { get; set; }
    }
}
