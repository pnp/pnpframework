using PnP.Framework.Migration.Pages.Ingredients;
using PnP.Framework.Migration.Taxonomy.Assets;
using System;
using System.Collections.Generic;

namespace PnP.Framework.Migration.Pages.Assessment
{
    /// <summary>
    /// Describes how far planning can advance for one captured ingredient before a
    /// target connection is available. This is evidence for planning; it is never
    /// executable mutation approval.
    /// </summary>
    public enum PageIngredientAssessmentState
    {
        Determined = 1,
        TargetInspectionRequired = 2,
        KnownGap = 3,
        AuthorizationBlocked = 4
    }

    public enum PageMigrationAssessmentState
    {
        ReadyForTargetInspection = 1,
        KnownGap = 2,
        AuthorizationBlocked = 3
    }

    /// <summary>
    /// Retained wire evidence for an ingredient request that returned literal
    /// HTTP 401 or 403. Redirects, CSOM payload errors, diagnostics, and inferred
    /// access-denied pages do not satisfy this contract.
    /// </summary>
    public sealed class PageIngredientAuthorizationEvidence
    {
        public string IngredientId { get; set; }

        public string Operation { get; set; }

        public string RequestUri { get; set; }

        public int HttpStatusCode { get; set; }

        public DateTimeOffset ObservedAtUtc { get; set; }

        public string EvidenceSource { get; set; }

        public string EvidenceSha256 { get; set; }
    }

    /// <summary>
    /// Additional retained evidence supplied to source-authoritative assessment.
    /// It is sealed into the resulting assessment digest and never grants mutation
    /// approval.
    /// </summary>
    public sealed class PageAssessmentEvidence
    {
        public string SchemaVersion { get; set; } = "pnp-page-assessment-evidence/v1";

        public IList<PageIngredientAuthorizationEvidence> AuthorizationFailures { get; set; } =
            new List<PageIngredientAuthorizationEvidence>();

        /// <summary>
        /// Optional digest-sealed, read-only target taxonomy preflight. Its
        /// deterministic mapping candidates may advance source-authoritative
        /// assessment, but they never grant taxonomy or page mutation approval.
        /// </summary>
        public TaxonomyAssetReviewPlan TaxonomyAssetReviewPlan { get; set; }
    }

    public sealed class PageIngredientAssessment
    {
        public string IngredientId { get; set; }

        public PageIngredientKind Kind { get; set; }

        public PageIngredientAssessmentState State { get; set; }

        public IngredientCapability Capability { get; set; }

        public IngredientDisposition ProposedDisposition { get; set; }

        public string ProposedRealization { get; set; }

        public string PolicyId { get; set; }

        public string Reason { get; set; }

        public string TargetIdentity { get; set; }

        public string MitigationCode { get; set; }

        public PageIngredientAuthorizationEvidence AuthorizationEvidence { get; set; }

        public IList<string> RequiredDependencyIngredientIds { get; set; } = new List<string>();

        public IList<string> VerificationAssertions { get; set; } = new List<string>();
    }
}
