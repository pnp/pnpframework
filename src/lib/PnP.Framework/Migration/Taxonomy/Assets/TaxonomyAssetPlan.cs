using PnP.Framework.Migration.Diagnostics;
using PnP.Framework.Migration.Evidence;
using System;
using System.Collections.Generic;
using System.Linq;

namespace PnP.Framework.Migration.Taxonomy.Assets
{
    public enum TaxonomyAssetTargetDisposition
    {
        TargetInspectionRequired = 1,
        CreateMissing = 2,
        ReuseOwned = 3,
        ReconcileOwnedPlanDrift = 4,
        ReviewExternalReuse = 5,
        CreateMissingAfterExternalApproval = 6,
        ResolveCollision = 7,
        AuthorizationBlocked = 8,
        RetryRequired = 9
    }

    public sealed class TaxonomyTermGroupMaterializationPlan
    {
        public string SchemaVersion { get; set; } = "pnp-taxonomy-termgroup-plan/v1";

        public TaxonomyTermGroupSourceIdentity Source { get; set; }

        public Guid TargetTermStoreId { get; set; }

        public Guid PreferredTargetGroupId { get; set; }

        public string TargetGroupName { get; set; }

        public string PlanDigest { get; set; }
    }

    public sealed class TaxonomyTermSetMaterializationPlan
    {
        public string SchemaVersion { get; set; } = "pnp-taxonomy-termset-plan/v1";

        public TaxonomyTermSetSourceIdentity Source { get; set; }

        public Guid TargetTermStoreId { get; set; }

        public Guid TargetGroupId { get; set; }

        public string TargetGroupName { get; set; }

        public Guid PreferredTargetTermSetId { get; set; }

        public string SourceTermSetName { get; set; }

        public string TargetTermSetName { get; set; }

        public int Language { get; set; }

        public bool IsOpenForTermCreation { get; set; }

        public bool IsAvailableForTagging { get; set; }

        public string OriginalIdentifierPropertyName { get; set; }

        public string OriginalIdentifier { get; set; }

        public string SourceEvidenceSha256 { get; set; }

        public string PlanDigest { get; set; }
    }

    public sealed class TaxonomyTermMaterializationPlan
    {
        public string SchemaVersion { get; set; } = "pnp-taxonomy-term-plan/v2";

        public TaxonomyTermSourceIdentity Source { get; set; }

        public Guid TargetTermStoreId { get; set; }

        public Guid TargetTermSetId { get; set; }

        public Guid? TargetParentTermId { get; set; }

        public Guid PreferredTargetTermId { get; set; }

        public string Name { get; set; }

        public string SourcePath { get; set; }

        public int Language { get; set; }

        public bool IsAvailableForTagging { get; set; }

        public bool? SourceIsReused { get; set; }

        public bool? SourceIsSourceTerm { get; set; }

        public Guid? SourceReuseSourceTermId { get; set; }

        public IList<Guid> SourceTermSetIds { get; set; } = new List<Guid>();

        public Guid? SourcePinSourceTermSetId { get; set; }

        public string OriginalIdentifierPropertyName { get; set; }

        public string OriginalIdentifier { get; set; }

        public string SourceEvidenceSha256 { get; set; }

        public string PlanDigest { get; set; }
    }

    public sealed class TaxonomyTermSetTargetProbe
    {
        public Guid SourceTermStoreId { get; set; }

        public Guid SourceTermSetId { get; set; }

        public Guid TargetTermStoreId { get; set; }

        public Guid? ResolvedTargetTermSetId { get; set; }

        public TaxonomyAssetTargetDisposition Disposition { get; set; }

        public IList<Guid> ProvenanceMatches { get; set; } = new List<Guid>();

        public bool PreferredIdExists { get; set; }

        public string ExistingName { get; set; }

        public string ExistingOriginalIdentifier { get; set; }

        public bool? ExistingIsOpenForTermCreation { get; set; }

        public bool? ExistingIsAvailableForTagging { get; set; }

        public bool ExternalCandidateInferredFromTerms { get; set; }

        public IList<Guid> SupportingTargetTermIds { get; set; } = new List<Guid>();

        public IList<Guid> MissingTargetTermIds { get; set; } = new List<Guid>();

        public LiteralHttpAuthorizationEvidence AuthorizationEvidence { get; set; }

        public IList<MigrationIssue> Issues { get; set; } = new List<MigrationIssue>();
    }

    public sealed class TaxonomyTermGroupTargetProbe
    {
        public Guid SourceTenantId { get; set; }

        public Guid SourceTermStoreId { get; set; }

        public Guid TargetTermStoreId { get; set; }

        public Guid? ResolvedTargetGroupId { get; set; }

        public TaxonomyAssetTargetDisposition Disposition { get; set; }

        public bool PreferredIdExists { get; set; }

        public string ExistingName { get; set; }

        public LiteralHttpAuthorizationEvidence AuthorizationEvidence { get; set; }

        public IList<MigrationIssue> Issues { get; set; } = new List<MigrationIssue>();
    }

    public sealed class TaxonomyTermTargetProbe
    {
        public Guid SourceTermStoreId { get; set; }

        public Guid SourceTermSetId { get; set; }

        public Guid SourceTermId { get; set; }

        public Guid TargetTermStoreId { get; set; }

        public Guid TargetTermSetId { get; set; }

        public Guid? ResolvedTargetTermId { get; set; }

        public TaxonomyAssetTargetDisposition Disposition { get; set; }

        public IList<Guid> ProvenanceMatches { get; set; } = new List<Guid>();

        public bool PreferredIdExists { get; set; }

        public string ExistingName { get; set; }

        public string ExistingPath { get; set; }

        public string ExistingOriginalIdentifier { get; set; }

        public Guid? ExistingParentTermId { get; set; }

        public Guid? ExistingTermSetId { get; set; }

        public bool? ExistingIsAvailableForTagging { get; set; }

        public bool? ExistingIsReused { get; set; }

        public bool? ExistingIsSourceTerm { get; set; }

        public Guid? ExistingReuseSourceTermId { get; set; }

        public IList<Guid> ExistingTermSetIds { get; set; } = new List<Guid>();

        public Guid? ExistingPinSourceTermSetId { get; set; }

        public LiteralHttpAuthorizationEvidence AuthorizationEvidence { get; set; }

        public IList<MigrationIssue> Issues { get; set; } = new List<MigrationIssue>();
    }

    public sealed class TaxonomyAssetMappingCandidate
    {
        public Guid SourceTermStoreId { get; set; }

        public Guid SourceTermSetId { get; set; }

        public Guid TargetTermStoreId { get; set; }

        public Guid TargetTermSetId { get; set; }

        public TaxonomyAssetTargetDisposition Disposition { get; set; }

        public bool RequiresReview { get; set; }

        public string EvidenceSha256 { get; set; }

        public IList<string> VerificationAssertions { get; set; } = new List<string>();
    }

    public sealed class TaxonomyAssetReviewPlan
    {
        public string SchemaVersion { get; set; } = "pnp-taxonomy-asset-review-plan/v1";

        public string SourceSnapshotDigest { get; set; }

        public Guid TargetTermStoreId { get; set; }

        public IList<TaxonomyTermGroupMaterializationPlan> TermGroups { get; set; } = new List<TaxonomyTermGroupMaterializationPlan>();

        public IList<TaxonomyTermSetMaterializationPlan> TermSets { get; set; } = new List<TaxonomyTermSetMaterializationPlan>();

        public IList<TaxonomyTermMaterializationPlan> Terms { get; set; } = new List<TaxonomyTermMaterializationPlan>();

        public IList<TaxonomyTermSetTargetProbe> TermSetProbes { get; set; } = new List<TaxonomyTermSetTargetProbe>();

        public IList<TaxonomyTermGroupTargetProbe> TermGroupProbes { get; set; } = new List<TaxonomyTermGroupTargetProbe>();

        public IList<TaxonomyTermTargetProbe> TermProbes { get; set; } = new List<TaxonomyTermTargetProbe>();

        public IList<TaxonomyAssetMappingCandidate> MappingCandidates { get; set; } = new List<TaxonomyAssetMappingCandidate>();

        public IList<MigrationIssue> Issues { get; set; } = new List<MigrationIssue>();

        public IList<LiteralHttpAuthorizationEvidence> AuthorizationStops { get; set; } = new List<LiteralHttpAuthorizationEvidence>();

        public string PlanDigest { get; set; }

        public bool HasAuthorizationStop => AuthorizationStops.Count > 0
            || TermGroupProbes.Any(value => value.Disposition == TaxonomyAssetTargetDisposition.AuthorizationBlocked)
            || TermSetProbes.Any(value => value.Disposition == TaxonomyAssetTargetDisposition.AuthorizationBlocked)
            || TermProbes.Any(value => value.Disposition == TaxonomyAssetTargetDisposition.AuthorizationBlocked);
    }
}
