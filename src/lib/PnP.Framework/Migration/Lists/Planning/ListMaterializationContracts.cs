using PnP.Framework.Migration.Diagnostics;
using PnP.Framework.Migration.Lists.Capture;
using PnP.Framework.Migration.Lists.Views;
using PnP.Framework.Migration.Topology;
using PnP.Framework.Migration.Schema.ContentTypes;
using PnP.Framework.Migration.Features;
using PnP.Framework.Migration.Packaging;
using System;
using System.Collections.Generic;
using System.Linq;

namespace PnP.Framework.Migration.Lists.Planning
{
    public enum ListMaterializationDisposition
    {
        CreateOwned = 1,
        ReuseOwned = 2,
        Block = 3
    }

    public enum ListFieldMaterializationDisposition
    {
        RequireTargetRuntime = 1,
        RequireTargetRuntimeAndCopyValue = 2,
        CreateOrReuseOwnedAndCopyValue = 3,
        CreateOrReuseOwnedCalculated = 4,
        MapLookup = 5,
        MapTaxonomy = 6,
        EvidenceOnly = 7,
        Block = 8,
        CreateOrReuseOwnedSchemaOnly = 9
    }

    public enum ListViewMaterializationDisposition
    {
        CreateOrReuseOwnedPublicView = 1,
        CreateOrReuseWebPartView = 2,
        SkipPersonal = 3,
        Block = 4
    }

    public enum ListViewRenderingResourceMaterializationDisposition
    {
        CreateOrReuseExact = 1,
        Block = 2,
        PreserveReferenceOnly = 3
    }

    public sealed class ListTargetOverride
    {
        public Guid SourceListId { get; set; }

        public string TargetTitle { get; set; }

        public string TargetRootFolderServerRelativeUrl { get; set; }
    }

    public sealed class ListFieldMaterializationPlan
    {
        public Guid SourceFieldId { get; set; }

        public string InternalName { get; set; }

        public string Title { get; set; }

        public string TypeAsString { get; set; }

        public ListFieldMaterializationDisposition Disposition { get; set; }

        public string SourceSchemaXml { get; set; }

        public string TargetSchemaXml { get; set; }

        public string SourcePortableSchemaSha256 { get; set; }

        public string TargetPortableSchemaSha256 { get; set; }

        public Guid? SourceLookupWebId { get; set; }

        public Guid? SourceLookupListId { get; set; }

        public string LookupField { get; set; }

        public string Reason { get; set; }
    }

    public sealed class ListViewMaterializationPlan
    {
        public Guid SourceViewId { get; set; }

        public string Title { get; set; }

        public string SourceServerRelativeUrl { get; set; }

        public ListViewMaterializationDisposition Disposition { get; set; }

        public ListViewSnapshot Source { get; set; }

        public string Reason { get; set; }
    }

    public sealed class ListViewRenderingResourceMaterializationPlan
    {
        public string SourceResourceId { get; set; }

        public ListViewRenderingResourceKind Kind { get; set; }

        public string SourceAbsoluteUrl { get; set; }

        public string SourceServerRelativeUrl { get; set; }

        public string TargetAbsoluteUrl { get; set; }

        public string TargetServerRelativeUrl { get; set; }

        public ArtifactReference SourceArtifact { get; set; }

        public string SourceContentBase64 { get; set; }

        public ListViewRenderingResourceMaterializationDisposition Disposition { get; set; }

        public string Reason { get; set; }

        public bool IsExecutable => Disposition != ListViewRenderingResourceMaterializationDisposition.Block;
    }

    public sealed class ListTargetProbe
    {
        public string PreferredTargetRootFolderServerRelativeUrl { get; set; }

        public string PreferredTargetTitle { get; set; }

        public string TargetWebUrl { get; set; }

        public string TargetRootFolderServerRelativeUrl { get; set; }

        public string TargetTitle { get; set; }

        public bool CollisionResolved { get; set; }

        public string CollisionResolutionReason { get; set; }

        public bool TargetWebExists { get; set; }

        public bool DeferredUntilTopologyMaterialization { get; set; }

        public Guid? TargetWebId { get; set; }

        public bool ListExists { get; set; }

        public Guid? TargetListId { get; set; }

        public string ExistingTitle { get; set; }

        public int? ExistingBaseTemplate { get; set; }

        public string ExistingOriginalIdentifier { get; set; }

        public string ExistingPlanDigest { get; set; }

        public IList<string> SameTitleDifferentPaths { get; set; } = new List<string>();

        public bool CanManageLists { get; set; }

        public ListMaterializationDisposition Disposition { get; set; }

        public IList<MigrationIssue> Issues { get; set; } = new List<MigrationIssue>();

        public bool IsAdmitted => Issues.All(value => value.Severity != MigrationIssueSeverity.Blocker && value.Severity != MigrationIssueSeverity.Error)
            && Disposition != ListMaterializationDisposition.Block;
    }

    public sealed class ListMaterializationPlan
    {
        public Guid SourceSiteId { get; set; }

        public Guid SourceWebId { get; set; }

        public Guid SourceListId { get; set; }

        public string TargetWebUrl { get; set; }

        public string TargetSiteCollectionUrl { get; set; }

        public string TargetWebServerRelativeUrl { get; set; }

        public string PreferredTargetRootFolderServerRelativeUrl { get; set; }

        public string TargetRootFolderServerRelativeUrl { get; set; }

        public string PreferredTargetTitle { get; set; }

        public string TargetTitle { get; set; }

        public string OriginalIdentifier { get; set; }

        public ListMaterializationDisposition Disposition { get; set; }

        public IList<ListFieldMaterializationPlan> Fields { get; set; } = new List<ListFieldMaterializationPlan>();

        public IList<ListViewMaterializationPlan> Views { get; set; } = new List<ListViewMaterializationPlan>();

        public IList<ListViewRenderingResourceMaterializationPlan> ViewRenderingResources { get; set; } = new List<ListViewRenderingResourceMaterializationPlan>();

        public IList<ContentTypeClosureNodePlan> SiteContentTypes { get; set; } = new List<ContentTypeClosureNodePlan>();

        public IList<PlatformFeatureMaterializationPlan> RequiredFeatures { get; set; } = new List<PlatformFeatureMaterializationPlan>();

        public IList<MigrationIssue> Issues { get; set; } = new List<MigrationIssue>();

        public ListTargetProbe TargetProbe { get; set; }

        public string PlanDigest { get; set; }

        public bool IsExecutable => Disposition != ListMaterializationDisposition.Block
            && Issues.All(value => value.Severity != MigrationIssueSeverity.Blocker && value.Severity != MigrationIssueSeverity.Error)
            && SiteContentTypes.All(value => value.IsExecutable)
            && RequiredFeatures.All(value => value.IsExecutable)
            && ViewRenderingResources.All(value => value.IsExecutable)
            && (TargetProbe == null || TargetProbe.IsAdmitted);
    }

    public sealed class ListMigrationPlanSet
    {
        public string SchemaVersion { get; set; } = "pnp-list-migration-plan/v1";

        public IList<Guid> OrderedSourceListIds { get; set; } = new List<Guid>();

        public IList<ListMaterializationPlan> Lists { get; set; } = new List<ListMaterializationPlan>();

        public IList<MigrationIssue> Issues { get; set; } = new List<MigrationIssue>();

        public string PlanDigest { get; set; }

        public bool IsExecutable => Lists.All(value => value.IsExecutable)
            && Issues.All(value => value.Severity != MigrationIssueSeverity.Blocker && value.Severity != MigrationIssueSeverity.Error);
    }

    public sealed class ListMaterializationReceipt
    {
        public Guid SourceWebId { get; set; }

        public Guid SourceListId { get; set; }

        public Guid TargetWebId { get; set; }

        public Guid TargetListId { get; set; }

        public string TargetRootFolderServerRelativeUrl { get; set; }

        public IDictionary<int, int> TargetItemIds { get; set; } = new Dictionary<int, int>();

        public IDictionary<Guid, Guid> TargetViewIds { get; set; } = new Dictionary<Guid, Guid>();

        public IDictionary<string, string> TargetContentTypeIds { get; set; } = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

        public ListMaterializationDisposition Disposition { get; set; }

        public string PlanDigest { get; set; }

        public bool FreshReadbackPassed { get; set; }

        public int VerifiedFieldCount { get; set; }

        public int VerifiedContentTypeCount { get; set; }

        public int VerifiedViewCount { get; set; }

        public int VerifiedViewRenderingResourceCount { get; set; }

        public int VerifiedItemCount { get; set; }

        public int VerifiedDocumentCount { get; set; }

        public int VerifiedAttachmentCount { get; set; }

        public IList<string> Diagnostics { get; set; } = new List<string>();
    }
}
