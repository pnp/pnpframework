using System;
using System.Collections.Generic;

namespace PnP.Framework.EnterpriseWiki
{
    public enum EnterpriseWikiPackageState
    {
        Draft = 0,
        ApprovalReady = 1,
        Blocked = 2
    }

    public enum EnterpriseWikiDependencyKind
    {
        Unknown = 0,
        Anchor = 1,
        Image = 2,
        Script = 3,
        StyleSheet = 4,
        IFrame = 5,
        Object = 6,
        Media = 7
    }

    public enum EnterpriseWikiDependencyDisposition
    {
        PreserveExternal = 0,
        RewriteToTarget = 1,
        MaterializeAtTarget = 2,
        Delegate = 3,
        Block = 4
    }

    public enum EnterpriseWikiFieldValueKind
    {
        String = 0,
        Boolean = 1,
        Number = 2,
        DateTime = 3,
        Url = 4,
        Taxonomy = 5,
        TaxonomyCollection = 6,
        Unsupported = 7
    }

    public sealed class EnterpriseWikiCaptureOptions
    {
        public string SourcePageServerRelativeUrl { get; set; }

        public string TargetPageServerRelativeUrl { get; set; }

        public bool IncludeWebParts { get; set; } = true;

        public bool Publish { get; set; } = true;

        public bool RequireInheritedPermissions { get; set; } = true;

        public bool BlockOnManagedMetadata { get; set; } = true;

        public bool AllowExternalResourceReferences { get; set; } = true;

        public long MaximumDependencyBytes { get; set; } = 10 * 1024 * 1024;
    }

    public sealed class EnterpriseWikiMigrationPackage
    {
        public string SchemaVersion { get; set; } = EnterpriseWikiPackageSerializer.SchemaVersion;

        public DateTimeOffset CreatedAtUtc { get; set; }

        public EnterpriseWikiPackageState State { get; set; }

        public EnterpriseWikiSnapshot Snapshot { get; set; }

        public EnterpriseWikiMigrationPlan Plan { get; set; }

        public string SnapshotDigest { get; set; }

        public string PlanDigest { get; set; }

        public EnterpriseWikiCustomerReport Report { get; set; }
    }

    public sealed class EnterpriseWikiSnapshot
    {
        public EnterpriseWikiPageIdentity Source { get; set; }

        public string PublishingPageContent { get; set; }

        public string PublishingPageContentSha256 { get; set; }

        public IList<EnterpriseWikiFieldValueSnapshot> Fields { get; set; } = new List<EnterpriseWikiFieldValueSnapshot>();

        public IList<EnterpriseWikiWebPartSnapshot> WebParts { get; set; } = new List<EnterpriseWikiWebPartSnapshot>();

        public IList<EnterpriseWikiDependencySnapshot> Dependencies { get; set; } = new List<EnterpriseWikiDependencySnapshot>();

        public EnterpriseWikiSecuritySnapshot Security { get; set; }

        public EnterpriseWikiLifecycleSnapshot Lifecycle { get; set; }

        public EnterpriseWikiSourceFence SourceFence { get; set; }
    }

    public sealed class EnterpriseWikiPageIdentity
    {
        public string WebUrl { get; set; }

        public string WebServerRelativeUrl { get; set; }

        public string PageServerRelativeUrl { get; set; }

        public int ListItemId { get; set; }

        public Guid FileUniqueId { get; set; }

        public string ContentTypeId { get; set; }

        public string ContentTypeName { get; set; }

        public string VersionLabel { get; set; }

        public long Length { get; set; }

        public DateTime ModifiedUtc { get; set; }

        public string Title { get; set; }

        public string PageLayoutUrl { get; set; }

        public string PageLayoutDescription { get; set; }
    }

    public sealed class EnterpriseWikiSourceFence
    {
        public Guid FileUniqueId { get; set; }

        public string VersionLabel { get; set; }

        public long Length { get; set; }

        public DateTime ModifiedUtc { get; set; }
    }

    public sealed class EnterpriseWikiFieldValueSnapshot
    {
        public string InternalName { get; set; }

        public string TypeAsString { get; set; }

        public EnterpriseWikiFieldValueKind Kind { get; set; }

        public string Value { get; set; }

        public bool ReadOnly { get; set; }

        public bool Hidden { get; set; }

        public bool Required { get; set; }
    }

    public sealed class EnterpriseWikiWebPartSnapshot
    {
        public Guid Id { get; set; }

        public string Title { get; set; }

        public string ZoneId { get; set; }

        public int ZoneIndex { get; set; }

        public bool Hidden { get; set; }

        public string ExportXml { get; set; }

        public string ExportSha256 { get; set; }
    }

    public sealed class EnterpriseWikiDependencySnapshot
    {
        public string OriginalValue { get; set; }

        public string SourceAbsoluteUrl { get; set; }

        public string SourceServerRelativeUrl { get; set; }

        public string TargetAbsoluteUrl { get; set; }

        public string TargetServerRelativeUrl { get; set; }

        public string Consumer { get; set; }

        public EnterpriseWikiDependencyKind Kind { get; set; }

        public EnterpriseWikiDependencyDisposition Disposition { get; set; }

        public string ContentBase64 { get; set; }

        public string ContentSha256 { get; set; }

        public long ContentLength { get; set; }

        public IList<string> Diagnostics { get; set; } = new List<string>();
    }

    public sealed class EnterpriseWikiSecuritySnapshot
    {
        public bool HasUniqueRoleAssignments { get; set; }

        public IList<EnterpriseWikiRoleAssignmentSnapshot> RoleAssignments { get; set; } = new List<EnterpriseWikiRoleAssignmentSnapshot>();
    }

    public sealed class EnterpriseWikiRoleAssignmentSnapshot
    {
        public string PrincipalLoginName { get; set; }

        public string PrincipalTitle { get; set; }

        public IList<string> RoleDefinitionNames { get; set; } = new List<string>();
    }

    public sealed class EnterpriseWikiLifecycleSnapshot
    {
        public string CheckOutType { get; set; }

        public string Level { get; set; }

        public int? ModerationStatus { get; set; }

        public DateTime CreatedUtc { get; set; }

        public DateTime ModifiedUtc { get; set; }
    }

    public sealed class EnterpriseWikiMigrationPlan
    {
        public string SourceSnapshotDigest { get; set; }

        public string SourceWebUrl { get; set; }

        public string SourcePageServerRelativeUrl { get; set; }

        public string TargetWebUrl { get; set; }

        public string TargetWebServerRelativeUrl { get; set; }

        public string TargetPageServerRelativeUrl { get; set; }

        public string PageLayoutName { get; set; }

        public bool Publish { get; set; }

        public bool CreateOnly { get; set; } = true;

        public EnterpriseWikiTargetProbe TargetProbe { get; set; }

        public IList<EnterpriseWikiTextReplacement> Replacements { get; set; } = new List<EnterpriseWikiTextReplacement>();

        public IList<string> StorageAssertions { get; set; } = new List<string>();

        public IList<string> BrowserAssertions { get; set; } = new List<string>();

        public IList<string> Blockers { get; set; } = new List<string>();

        public IList<string> Warnings { get; set; } = new List<string>();

        public bool IsExecutable => Blockers.Count == 0;
    }

    public sealed class EnterpriseWikiTextReplacement
    {
        public string Source { get; set; }

        public string Target { get; set; }

        public string Reason { get; set; }
    }

    public sealed class EnterpriseWikiTargetProbe
    {
        public string WebUrl { get; set; }

        public string WebServerRelativeUrl { get; set; }

        public string WebTemplate { get; set; }

        public int WebConfiguration { get; set; }

        public string PagesLibraryServerRelativeUrl { get; set; }

        public int PagesLibraryBaseTemplate { get; set; }

        public string EnterpriseWikiContentTypeId { get; set; }

        public string EnterpriseWikiLayoutUrl { get; set; }

        public bool EnterpriseWikiLayoutExists { get; set; }

        public bool TargetPageExists { get; set; }

        public IList<string> ExistingDependencyPaths { get; set; } = new List<string>();
    }

    public sealed class EnterpriseWikiCustomerReport
    {
        public string Summary { get; set; }

        public IList<string> CapturedIngredients { get; set; } = new List<string>();

        public IList<string> Blockers { get; set; } = new List<string>();

        public IList<string> Warnings { get; set; } = new List<string>();
    }

    public sealed class EnterpriseWikiCopyReceipt
    {
        public string SchemaVersion { get; set; } = "pnp-enterprise-wiki-copy-receipt/v1";

        public DateTimeOffset StartedAtUtc { get; set; }

        public DateTimeOffset CompletedAtUtc { get; set; }

        public string ApprovedPlanDigest { get; set; }

        public string TargetWebUrl { get; set; }

        public string TargetPageServerRelativeUrl { get; set; }

        public Guid TargetFileUniqueId { get; set; }

        public int TargetListItemId { get; set; }

        public string TargetContentTypeId { get; set; }

        public string TargetVersionLabel { get; set; }

        public string PersistedPublishingPageContentSha256 { get; set; }

        public bool StorageContentEqual { get; set; }

        public int ImportedWebPartCount { get; set; }

        public int MaterializedDependencyCount { get; set; }

        public bool FreshReadbackPassed { get; set; }

        public IList<string> Warnings { get; set; } = new List<string>();

        public bool Succeeded { get; set; }
    }
}
