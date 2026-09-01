using System;
using System.Collections.Generic;

namespace PnP.Framework.Migration.Execution
{
    public sealed class MigrationMutationReceipt
    {
        public Guid OperationId { get; set; }

        public string PlanDigest { get; set; }

        public string ActionId { get; set; }

        public int Sequence { get; set; }

        public DateTimeOffset CompletedAtUtc { get; set; }

        public MutationOutcome Outcome { get; set; }

        public IList<string> ExchangeIds { get; set; } = new List<string>();

        public string Message { get; set; }
    }
}
