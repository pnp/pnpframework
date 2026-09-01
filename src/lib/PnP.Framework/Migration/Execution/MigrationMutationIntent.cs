using System;

namespace PnP.Framework.Migration.Execution
{
    public sealed class MigrationMutationIntent
    {
        public Guid OperationId { get; set; }

        public string PlanDigest { get; set; }

        public string ActionId { get; set; }

        public int Sequence { get; set; }

        public DateTimeOffset WrittenAtUtc { get; set; }

        public string Description { get; set; }
    }
}
