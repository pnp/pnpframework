using PnP.Framework.Migration.Evidence;
using System;
using System.Collections.Generic;

namespace PnP.Framework.Migration.Taxonomy.Assets
{
    /// <summary>
    /// Describes the source taxonomy dependency closure required by one or more
    /// captured fields, anchors, or values. It intentionally does not imply that
    /// the complete source TermSet must be copied.
    /// </summary>
    public sealed class TaxonomyTermSetCaptureRequest
    {
        public Guid SourceTermStoreId { get; set; }

        public Guid SourceTermSetId { get; set; }

        public IList<string> SourceWebUrls { get; set; } = new List<string>();

        public IList<Guid> RequiredTermIds { get; set; } = new List<Guid>();

        public IList<string> Consumers { get; set; } = new List<string>();
    }

    public sealed class TaxonomyTermSetSourceSnapshot
    {
        public string SchemaVersion { get; set; } = "pnp-taxonomy-termset-source/v1";

        public Guid SourceTenantId { get; set; }

        public Guid SourceTermStoreId { get; set; }

        public Guid SourceTermSetId { get; set; }

        public string SourceWebUrl { get; set; }

        public string Name { get; set; }

        public int Language { get; set; } = 1033;

        public bool IsOpenForTermCreation { get; set; }

        public bool IsAvailableForTagging { get; set; }

        public EvidenceAvailability Availability { get; set; } = EvidenceAvailability.Captured;

        public LiteralHttpAuthorizationEvidence AuthorizationEvidence { get; set; }

        public IList<string> Consumers { get; set; } = new List<string>();

        public IList<string> Diagnostics { get; set; } = new List<string>();

        public string EvidenceSha256 { get; set; }
    }

    public sealed class TaxonomyTermSourceSnapshot
    {
        public string SchemaVersion { get; set; } = "pnp-taxonomy-term-source/v2";

        public Guid SourceTenantId { get; set; }

        public Guid SourceTermStoreId { get; set; }

        public Guid SourceTermSetId { get; set; }

        public Guid SourceTermId { get; set; }

        public string SourceWebUrl { get; set; }

        public Guid? SourceParentTermId { get; set; }

        public string Name { get; set; }

        public string Path { get; set; }

        public int Language { get; set; } = 1033;

        public bool IsAvailableForTagging { get; set; }

        /// <summary>
        /// Captures whether SharePoint reports this Term instance as reused.
        /// Null means the Term was unavailable and the relationship was not observed.
        /// </summary>
        public bool? IsReused { get; set; }

        /// <summary>
        /// Captures whether SharePoint reports this Term as the source Term for reuse.
        /// Null means the Term was unavailable and the relationship was not observed.
        /// </summary>
        public bool? IsSourceTerm { get; set; }

        /// <summary>
        /// Gets the source Term identity reported for a reused Term, when present.
        /// </summary>
        public Guid? ReuseSourceTermId { get; set; }

        /// <summary>
        /// Gets every TermSet membership reported for the Term identity. This is
        /// retained because one Term GUID can participate in multiple TermSets.
        /// </summary>
        public IList<Guid> TermSetIds { get; set; } = new List<Guid>();

        /// <summary>
        /// Gets the source TermSet used for a pinned Term relationship, when present.
        /// </summary>
        public Guid? PinSourceTermSetId { get; set; }

        public EvidenceAvailability Availability { get; set; } = EvidenceAvailability.Captured;

        public LiteralHttpAuthorizationEvidence AuthorizationEvidence { get; set; }

        public IList<string> Diagnostics { get; set; } = new List<string>();

        public string EvidenceSha256 { get; set; }
    }

    public sealed class TaxonomyAssetSourceSnapshot
    {
        public string SchemaVersion { get; set; } = "pnp-taxonomy-asset-source/v2";

        public Guid SourceTenantId { get; set; }

        public IList<TaxonomyTermSetSourceSnapshot> TermSets { get; set; } = new List<TaxonomyTermSetSourceSnapshot>();

        public IList<TaxonomyTermSourceSnapshot> Terms { get; set; } = new List<TaxonomyTermSourceSnapshot>();

        public IList<string> Diagnostics { get; set; } = new List<string>();

        public string SnapshotDigest { get; set; }
    }
}
