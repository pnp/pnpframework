using System;
using System.Collections.Generic;

namespace PnP.Framework.Migration.Verification
{
    public sealed class RuntimeVerificationReceipt
    {
        public string SchemaVersion { get; set; } = "pnp-migration-runtime-verification-receipt/v1";

        public string PlanDigest { get; set; }

        public string TargetIdentity { get; set; }

        public DateTimeOffset CompletedAtUtc { get; set; }

        public IList<RuntimeVerificationResult> Results { get; set; } = new List<RuntimeVerificationResult>();

        public RuntimeVerificationStatus Status { get; set; }
    }
}
