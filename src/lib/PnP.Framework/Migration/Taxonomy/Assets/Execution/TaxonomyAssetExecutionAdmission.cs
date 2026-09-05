using PnP.Framework.Migration.Execution;
using System.Collections.Generic;

namespace PnP.Framework.Migration.Taxonomy.Assets.Execution
{
    public sealed class TaxonomyAssetExecutionAdmission
    {
        public string SchemaVersion { get; set; } = "pnp-taxonomy-asset-execution-admission/v1";

        public string ReviewPlanDigest { get; set; }

        public string ApprovalDigest { get; set; }

        public string FreshInspectionPlanDigest { get; set; }

        public bool IsAdmitted { get; set; }

        public IList<string> ApprovedActionIds { get; set; } = new List<string>();

        public IList<string> DeferredActionIds { get; set; } = new List<string>();

        public IList<string> RejectedActionIds { get; set; } = new List<string>();

        public IList<ExecutionAdmissionFailure> Failures { get; set; } = new List<ExecutionAdmissionFailure>();
    }
}
