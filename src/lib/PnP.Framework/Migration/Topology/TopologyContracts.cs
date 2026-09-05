using PnP.Framework.Migration.Diagnostics;
using PnP.Framework.Migration.Evidence;
using System;
using System.Collections.Generic;
using System.Linq;

namespace PnP.Framework.Migration.Topology
{
    public enum TopologyNodeKind
    {
        SiteCollectionRoot = 1,
        ChildWeb = 2
    }

    public enum TargetSiteMode
    {
        ExistingTargetSite = 1,
        CreateTargetSite = 2
    }

    public enum TopologyMaterializationDisposition
    {
        CreateOwned = 1,
        ReuseOwned = 2,
        ReuseApprovedHost = 3,
        Block = 4,
        RecoverInterruptedCreate = 5
    }

    public sealed class SourceWebSnapshot
    {
        public Guid SiteId { get; set; }

        public Guid WebId { get; set; }

        public Guid? ParentWebId { get; set; }

        public string SiteCollectionUrl { get; set; }

        public string WebUrl { get; set; }

        public string ServerRelativeUrl { get; set; }

        public string Title { get; set; }

        public string WebTemplate { get; set; }

        public int Configuration { get; set; }

        public EvidenceAvailability Availability { get; set; } = EvidenceAvailability.Captured;

        public IList<string> Diagnostics { get; set; } = new List<string>();
    }

    public sealed class SourceSiteCollectionSnapshot
    {
        public string SchemaVersion { get; set; } = "pnp-source-topology/v1";

        public Guid SiteId { get; set; }

        public string SiteCollectionUrl { get; set; }

        public string ServerRelativeUrl { get; set; }

        public Guid RootWebId { get; set; }

        public IList<SourceWebSnapshot> Webs { get; set; } = new List<SourceWebSnapshot>();

        public EvidenceAvailability Availability { get; set; } = EvidenceAvailability.Captured;

        public IList<string> Diagnostics { get; set; } = new List<string>();
    }

    public sealed class TargetSiteCollectionSpec
    {
        public Guid SourceSiteId { get; set; }

        public TargetSiteMode Mode { get; set; } = TargetSiteMode.ExistingTargetSite;

        public string TargetSiteUrl { get; set; }

        public Guid? ExpectedTargetSiteId { get; set; }

        public string Title { get; set; }

        public string Owner { get; set; }

        public string Template { get; set; } = "STS#3";

        public int Language { get; set; } = 1033;

        public int TimeZone { get; set; }
    }

    public sealed class TargetWebOverride
    {
        public Guid SourceWebId { get; set; }

        public string TargetUrlSegment { get; set; }

        public string TargetTitle { get; set; }

        public string TargetTemplate { get; set; }

        public int? TargetConfiguration { get; set; }
    }

    public sealed class TopologyPlanningPolicy
    {
        public bool PreserveSourceChildWebTemplate { get; set; } = true;

        public string DefaultChildWebTemplate { get; set; } = "STS#0";

        public int DefaultChildWebConfiguration { get; set; }

        public IList<TargetWebOverride> WebOverrides { get; set; } = new List<TargetWebOverride>();
    }

    public sealed class WebMappingPlan
    {
        public TopologyNodeKind Kind { get; set; }

        public Guid SourceSiteId { get; set; }

        public Guid SourceWebId { get; set; }

        public Guid? SourceParentWebId { get; set; }

        public string SourceSiteCollectionUrl { get; set; }

        public string SourceWebUrl { get; set; }

        public string SourceServerRelativeUrl { get; set; }

        public string TargetSiteCollectionUrl { get; set; }

        public string PreferredTargetWebUrl { get; set; }

        public string TargetWebUrl { get; set; }

        public string PreferredTargetServerRelativeUrl { get; set; }

        public string TargetServerRelativeUrl { get; set; }

        public string TargetParentWebUrl { get; set; }

        public string TargetTitle { get; set; }

        public string TargetTemplate { get; set; }

        public int TargetConfiguration { get; set; }

        public string OriginalIdentifier { get; set; }
    }

    public sealed class SiteCollectionMappingPlan
    {
        public Guid SourceSiteId { get; set; }

        public string SourceSiteCollectionUrl { get; set; }

        public TargetSiteMode TargetMode { get; set; }

        public string PreferredTargetSiteCollectionUrl { get; set; }

        public string TargetSiteCollectionUrl { get; set; }

        public bool TargetSiteCollisionResolved { get; set; }

        public string TargetSiteResolutionReason { get; set; }

        public Guid? ExpectedTargetSiteId { get; set; }

        public string TargetTitle { get; set; }

        public string TargetOwner { get; set; }

        public string TargetTemplate { get; set; }

        public int TargetLanguage { get; set; }

        public int TargetTimeZone { get; set; }

        public string OriginalIdentifier { get; set; }

        public IList<WebMappingPlan> Webs { get; set; } = new List<WebMappingPlan>();
    }

    public sealed class TopologyPlan
    {
        public string SchemaVersion { get; set; } = "pnp-topology-plan/v1";

        public IList<SiteCollectionMappingPlan> SiteCollections { get; set; } = new List<SiteCollectionMappingPlan>();

        public string PlanDigest { get; set; }
    }

    public sealed class TopologyPlanBuildResult
    {
        public TopologyPlan Plan { get; set; }

        public IList<MigrationIssue> Issues { get; set; } = new List<MigrationIssue>();

        public bool IsExecutable => Plan != null && Issues.All(issue => issue.Severity != MigrationIssueSeverity.Blocker && issue.Severity != MigrationIssueSeverity.Error);
    }

    public sealed class TopologyWebTargetProbe
    {
        public Guid SourceSiteId { get; set; }

        public Guid SourceWebId { get; set; }

        public string PreferredTargetWebUrl { get; set; }

        public string TargetWebUrl { get; set; }

        public string PreferredTargetServerRelativeUrl { get; set; }

        public string TargetServerRelativeUrl { get; set; }

        public bool CollisionResolved { get; set; }

        public string CollisionResolutionReason { get; set; }

        public bool Exists { get; set; }

        public Guid? TargetSiteId { get; set; }

        public Guid? TargetWebId { get; set; }

        public Guid? TargetParentWebId { get; set; }

        public string ExistingTitle { get; set; }

        public string ExistingTemplate { get; set; }

        public int? ExistingConfiguration { get; set; }

        public string ExistingOriginalIdentifier { get; set; }

        public string ExistingPlanDigest { get; set; }

        public TopologyMaterializationDisposition Disposition { get; set; }

        public IList<MigrationIssue> Issues { get; set; } = new List<MigrationIssue>();

        public bool IsAdmitted => Disposition != TopologyMaterializationDisposition.Block
            && Issues.All(value => value.Severity != MigrationIssueSeverity.Blocker && value.Severity != MigrationIssueSeverity.Error);
    }

    public sealed class TopologySiteTargetProbe
    {
        public Guid SourceSiteId { get; set; }

        public string PreferredTargetSiteCollectionUrl { get; set; }

        public string TargetSiteCollectionUrl { get; set; }

        public bool CollisionResolved { get; set; }

        public string CollisionResolutionReason { get; set; }

        public bool Exists { get; set; }

        public Guid? TargetSiteId { get; set; }

        public Guid? TargetRootWebId { get; set; }

        public TopologyMaterializationDisposition Disposition { get; set; }

        public IList<TopologyWebTargetProbe> Webs { get; set; } = new List<TopologyWebTargetProbe>();

        public IList<MigrationIssue> Issues { get; set; } = new List<MigrationIssue>();

        public bool IsAdmitted => Disposition != TopologyMaterializationDisposition.Block
            && Issues.All(value => value.Severity != MigrationIssueSeverity.Blocker && value.Severity != MigrationIssueSeverity.Error)
            && Webs.All(value => value.IsAdmitted);
    }

    public sealed class TopologyTargetAnalysis
    {
        public string SchemaVersion { get; set; } = "pnp-topology-target-analysis/v1";

        public string TopologyPlanDigest { get; set; }

        public IList<TopologySiteTargetProbe> SiteCollections { get; set; } = new List<TopologySiteTargetProbe>();

        public IList<MigrationIssue> Issues { get; set; } = new List<MigrationIssue>();

        public bool IsAdmitted => Issues.All(value => value.Severity != MigrationIssueSeverity.Blocker && value.Severity != MigrationIssueSeverity.Error)
            && SiteCollections.All(value => value.IsAdmitted);
    }

    public sealed class TopologyWebMaterializationReceipt
    {
        public Guid SourceSiteId { get; set; }

        public Guid SourceWebId { get; set; }

        public Guid TargetSiteId { get; set; }

        public Guid TargetWebId { get; set; }

        public string TargetWebUrl { get; set; }

        public TopologyMaterializationDisposition Disposition { get; set; }

        public string MappingDigest { get; set; }
    }

    public sealed class TopologyMaterializationReceipt
    {
        public string TopologyPlanDigest { get; set; }

        public string ApprovedTopologyPlanDigest { get; set; }

        public string ExecutionTopologyPlanDigest { get; set; }

        public IList<TopologyWebMaterializationReceipt> Webs { get; set; } = new List<TopologyWebMaterializationReceipt>();

        public bool FreshReadbackPassed { get; set; }

        public IList<string> Diagnostics { get; set; } = new List<string>();
    }
}
