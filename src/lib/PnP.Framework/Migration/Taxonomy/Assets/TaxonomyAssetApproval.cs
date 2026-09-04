using System;
using System.Collections.Generic;

namespace PnP.Framework.Migration.Taxonomy.Assets
{
    public enum TaxonomyAssetKind
    {
        TermSet = 1,
        Term = 2,
        TermGroup = 3
    }

    public enum TaxonomyAssetApprovalDecision
    {
        Pending = 1,
        Approve = 2,
        Defer = 3,
        Reject = 4
    }

    /// <summary>
    /// Records the human decision for one reviewed taxonomy asset action. The
    /// reviewed disposition and target identities are repeated for readability;
    /// the review-plan digest remains the authoritative approval boundary.
    /// </summary>
    public sealed class TaxonomyAssetActionApproval
    {
        public string ActionId { get; set; }

        public TaxonomyAssetKind Kind { get; set; }

        public Guid SourceTenantId { get; set; }

        public Guid SourceTermStoreId { get; set; }

        public Guid SourceTermSetId { get; set; }

        public Guid? SourceTermId { get; set; }

        public Guid TargetTermStoreId { get; set; }

        public Guid? TargetTermGroupId { get; set; }

        public Guid TargetTermSetId { get; set; }

        public Guid? TargetTermId { get; set; }

        public TaxonomyAssetTargetDisposition ReviewedDisposition { get; set; }

        public TaxonomyAssetApprovalDecision Decision { get; set; }

        public bool RequiresExplicitReview { get; set; }

        /// <summary>
        /// Must be true before an approved action may add a child asset to an
        /// external TermSet. It is never inferred from a generic Approve value.
        /// </summary>
        public bool ExternalMutationApproved { get; set; }

        public string Comment { get; set; }
    }

    public sealed class TaxonomyAssetApprovalManifest
    {
        public string SchemaVersion { get; set; } = "pnp-taxonomy-asset-approval/v1";

        public string ReviewPlanDigest { get; set; }

        public DateTimeOffset GeneratedAtUtc { get; set; }

        public DateTimeOffset? ApprovedAtUtc { get; set; }

        public string ApprovedBy { get; set; }

        public IList<TaxonomyAssetActionApproval> Actions { get; set; } = new List<TaxonomyAssetActionApproval>();

        public string ApprovalDigest { get; set; }
    }
}
