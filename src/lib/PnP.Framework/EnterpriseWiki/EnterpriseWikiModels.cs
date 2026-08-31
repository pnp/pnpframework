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

    public enum EnterpriseWikiCaptureStatus
    {
        Captured = 0,
        NotReturned = 1,
        CapturedWithLimitations = 2,
        Failed = 3
    }

    public enum EnterpriseWikiFieldValueKind
    {
        Null = 0,
        String = 1,
        StringCollection = 2,
        Boolean = 3,
        Number = 4,
        DateTime = 5,
        Guid = 6,
        Url = 7,
        User = 8,
        UserCollection = 9,
        Lookup = 10,
        LookupCollection = 11,
        Taxonomy = 12,
        TaxonomyCollection = 13,
        ByteArray = 14,
        Unsupported = 15
    }

    public enum EnterpriseWikiFieldDisposition
    {
        Apply = 0,
        AlreadyHandled = 1,
        SkipEmpty = 2,
        SkipReadOnly = 3,
        SkipCalculated = 4,
        TargetFieldMissing = 5,
        TargetTypeMismatch = 6,
        RequiresMapping = 7,
        EvidenceOnly = 8,
        CaptureUnavailable = 9,
        Block = 10
    }

    public enum EnterpriseWikiTargetLifecycle
    {
        Draft = 0,
        Published = 1
    }

    public enum EnterpriseWikiMigrationOperation
    {
        CreatePage = 0,
        ApplyDeferredFields = 1
    }

    public sealed class EnterpriseWikiExportOptions
    {
        public string SourcePageServerRelativeUrl { get; set; }

        public bool IncludeWebParts { get; set; } = true;

        public long MaximumDependencyBytes { get; set; } = 10 * 1024 * 1024;
    }

    public sealed class EnterpriseWikiPlanningOptions
    {
        public string TargetPageServerRelativeUrl { get; set; }

        public bool RequireInheritedPermissions { get; set; } = true;

        public bool BlockOnManagedMetadata { get; set; } = true;

        public bool AllowExternalResourceReferences { get; set; } = true;

        public bool CreateOnly { get; set; } = true;
    }

    public sealed class EnterpriseWikiExportPackage
    {
        public string SchemaVersion { get; set; } = EnterpriseWikiPackageSerializer.ExportSchemaVersion;

        public DateTimeOffset ExportedAtUtc { get; set; }

        public EnterpriseWikiSnapshot Snapshot { get; set; }

        public string SnapshotDigest { get; set; }
    }

    public sealed class EnterpriseWikiMigrationPackage
    {
        public string SchemaVersion { get; set; } = EnterpriseWikiPackageSerializer.MigrationSchemaVersion;

        public DateTimeOffset PlannedAtUtc { get; set; }

        public string ExportSchemaVersion { get; set; } = EnterpriseWikiPackageSerializer.ExportSchemaVersion;

        public DateTimeOffset ExportedAtUtc { get; set; }

        public EnterpriseWikiPackageState State { get; set; }

        public EnterpriseWikiSnapshot Snapshot { get; set; }

        public EnterpriseWikiMigrationPlan Plan { get; set; }

        public string SnapshotDigest { get; set; }

        public string PlanDigest { get; set; }

        public EnterpriseWikiCustomerReport Report { get; set; }
    }

    public sealed class EnterpriseWikiSnapshot
    {
        public EnterpriseWikiExportOptions CapturePolicy { get; set; }

        public EnterpriseWikiPageIdentity Source { get; set; }

        public string PublishingPageContent { get; set; }

        public string PublishingPageContentSha256 { get; set; }

        public IList<EnterpriseWikiFieldValueSnapshot> Fields { get; set; } = new List<EnterpriseWikiFieldValueSnapshot>();

        public IList<EnterpriseWikiWebPartSnapshot> WebParts { get; set; } = new List<EnterpriseWikiWebPartSnapshot>();

        public IList<EnterpriseWikiDependencySnapshot> Dependencies { get; set; } = new List<EnterpriseWikiDependencySnapshot>();

        public EnterpriseWikiSecuritySnapshot Security { get; set; }

        public EnterpriseWikiLifecycleSnapshot Lifecycle { get; set; }

        public EnterpriseWikiSourceFence SourceFence { get; set; }

        public IList<string> Blockers { get; set; } = new List<string>();

        public IList<string> Warnings { get; set; } = new List<string>();
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
        public Guid Id { get; set; }

        public string InternalName { get; set; }

        public string Title { get; set; }

        public string TypeAsString { get; set; }

        public string SchemaXml { get; set; }

        public bool ReadOnly { get; set; }

        public bool Hidden { get; set; }

        public bool Required { get; set; }

        public bool HasValue { get; set; }

        public EnterpriseWikiFieldValueKind Kind { get; set; }

        public string Value { get; set; }

        public IList<string> StringValues { get; set; } = new List<string>();

        public EnterpriseWikiUrlValueSnapshot UrlValue { get; set; }

        public IList<EnterpriseWikiLookupValueSnapshot> LookupValues { get; set; } = new List<EnterpriseWikiLookupValueSnapshot>();

        public IList<EnterpriseWikiTaxonomyValueSnapshot> TaxonomyValues { get; set; } = new List<EnterpriseWikiTaxonomyValueSnapshot>();

        public string BinaryBase64 { get; set; }

        public string RawType { get; set; }

        public string RawValue { get; set; }

        public string RawValueJson { get; set; }

        public EnterpriseWikiCaptureStatus CaptureStatus { get; set; }

        public IList<string> Diagnostics { get; set; } = new List<string>();
    }

    public sealed class EnterpriseWikiUrlValueSnapshot
    {
        public string Url { get; set; }

        public string Description { get; set; }
    }

    public sealed class EnterpriseWikiLookupValueSnapshot
    {
        public int LookupId { get; set; }

        public string LookupValue { get; set; }
    }

    public sealed class EnterpriseWikiTaxonomyValueSnapshot
    {
        public string Label { get; set; }

        public string TermGuid { get; set; }

        public int WssId { get; set; }
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
        public string Id { get; set; }

        public string OriginalValue { get; set; }

        public string SourceAbsoluteUrl { get; set; }

        public string SourceServerRelativeUrl { get; set; }

        public string Consumer { get; set; }

        public EnterpriseWikiDependencyKind Kind { get; set; }

        public bool IsRenderableResource { get; set; }

        public string ContentBase64 { get; set; }

        public string ContentSha256 { get; set; }

        public long ContentLength { get; set; }

        public EnterpriseWikiCaptureStatus CaptureStatus { get; set; }

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

        public EnterpriseWikiMigrationOperation Operation { get; set; } = EnterpriseWikiMigrationOperation.CreatePage;

        public EnterpriseWikiTargetLifecycle TargetLifecycle { get; set; }

        public string LifecycleReason { get; set; }

        public bool CreateOnly { get; set; } = true;

        public EnterpriseWikiPlanningOptions PlanningPolicy { get; set; }

        public EnterpriseWikiTargetProbe TargetProbe { get; set; }

        public IList<EnterpriseWikiFieldAction> FieldActions { get; set; } = new List<EnterpriseWikiFieldAction>();

        public IList<EnterpriseWikiDependencyAction> DependencyActions { get; set; } = new List<EnterpriseWikiDependencyAction>();

        public IList<EnterpriseWikiTextReplacement> Replacements { get; set; } = new List<EnterpriseWikiTextReplacement>();

        public string ExpectedPublishingPageContentSha256 { get; set; }

        public IList<string> StorageAssertions { get; set; } = new List<string>();

        public IList<string> BrowserAssertions { get; set; } = new List<string>();

        public IList<string> Blockers { get; set; } = new List<string>();

        public IList<string> Warnings { get; set; } = new List<string>();

        public bool IsExecutable => Blockers.Count == 0;
    }

    public sealed class EnterpriseWikiFieldAction
    {
        public string SourceInternalName { get; set; }

        public string TargetInternalName { get; set; }

        public string TargetTypeAsString { get; set; }

        public EnterpriseWikiFieldDisposition Disposition { get; set; }

        public string Reason { get; set; }

        public bool WillApply => Disposition == EnterpriseWikiFieldDisposition.Apply;
    }

    public sealed class EnterpriseWikiDependencyAction
    {
        public string SnapshotDependencyId { get; set; }

        public string TargetAbsoluteUrl { get; set; }

        public string TargetServerRelativeUrl { get; set; }

        public EnterpriseWikiDependencyDisposition Disposition { get; set; }

        public IList<string> Diagnostics { get; set; } = new List<string>();
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

        public bool EnableVersioning { get; set; }

        public bool EnableMinorVersions { get; set; }

        public bool EnableModeration { get; set; }

        public bool ForceCheckout { get; set; }

        public string DraftVersionVisibility { get; set; }

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

    public sealed class EnterpriseWikiFieldImportResult
    {
        public string InternalName { get; set; }

        public EnterpriseWikiFieldDisposition PlannedDisposition { get; set; }

        public bool Attempted { get; set; }

        public bool Succeeded { get; set; }

        public string Message { get; set; }
    }

    public sealed class EnterpriseWikiImportReceipt
    {
        public string SchemaVersion { get; set; } = "pnp-enterprise-wiki-import-receipt/v2";

        public DateTimeOffset StartedAtUtc { get; set; }

        public DateTimeOffset CompletedAtUtc { get; set; }

        public string ApprovedPlanDigest { get; set; }

        public string TargetWebUrl { get; set; }

        public string TargetPageServerRelativeUrl { get; set; }

        public Guid TargetFileUniqueId { get; set; }

        public int TargetListItemId { get; set; }

        public string TargetContentTypeId { get; set; }

        public string TargetVersionLabel { get; set; }

        public EnterpriseWikiTargetLifecycle ExpectedLifecycle { get; set; }

        public string ActualFileLevel { get; set; }

        public string ActualCheckOutType { get; set; }

        public int? ActualModerationStatus { get; set; }

        public bool LifecycleMatched { get; set; }

        public string ExpectedPublishingPageContentSha256 { get; set; }

        public string PersistedPublishingPageContentSha256 { get; set; }

        public bool StorageContentEqual { get; set; }

        public int ImportedWebPartCount { get; set; }

        public int MaterializedDependencyCount { get; set; }

        public IList<EnterpriseWikiFieldImportResult> FieldResults { get; set; } = new List<EnterpriseWikiFieldImportResult>();

        public bool FreshReadbackPassed { get; set; }

        public IList<string> Warnings { get; set; } = new List<string>();

        public bool Succeeded { get; set; }
    }
}
