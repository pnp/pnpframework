using PnP.Framework.Migration.Evidence;
using PnP.Framework.Migration.Lists.ContentTypes;
using PnP.Framework.Migration.Lists.Fields;
using PnP.Framework.Migration.Lists.Items;
using PnP.Framework.Migration.Lists.Views;
using PnP.Framework.Migration.Schema.ContentTypes;
using System;
using System.Collections.Generic;
using System.Text.Json.Serialization;

namespace PnP.Framework.Migration.Lists.Capture
{
    /// <summary>
    /// Captures the source List Information Rights Management policy separately
    /// from document bytes. SharePoint can generate an IRM download envelope at
    /// read time, so binary representation alone cannot prove whether the file
    /// itself or the owning library supplied the protection boundary.
    /// </summary>
    public sealed class ListInformationRightsManagementSnapshot
    {
        public bool IrmEnabled { get; set; }

        public bool IrmExpire { get; set; }

        public bool IrmReject { get; set; }

        [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
        public ListInformationRightsManagementPolicySnapshot Policy { get; set; }

        public EvidenceAvailability Availability { get; set; } = EvidenceAvailability.Captured;

        [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
        public LiteralHttpAuthorizationEvidence AuthorizationEvidence { get; set; }

        public IList<string> Diagnostics { get; set; } = new List<string>();
    }

    public sealed class ListInformationRightsManagementPolicySnapshot
    {
        public bool AllowPrint { get; set; }

        public bool AllowScript { get; set; }

        public bool AllowWriteCopy { get; set; }

        public bool DisableDocumentBrowserView { get; set; }

        public int DocumentAccessExpireDays { get; set; }

        public DateTime DocumentLibraryProtectionExpireDate { get; set; }

        public bool EnableDocumentAccessExpire { get; set; }

        public bool EnableDocumentBrowserPublishingView { get; set; }

        public bool EnableGroupProtection { get; set; }

        public bool EnableLicenseCacheExpire { get; set; }

        public string GroupName { get; set; }

        public int LicenseCacheExpireDays { get; set; }

        public string PolicyDescription { get; set; }

        public string PolicyTitle { get; set; }

        public string TemplateId { get; set; }
    }

    public sealed class ListDependencySnapshot
    {
        public string SchemaVersion { get; set; } = "pnp-list-dependency/v1";

        public Guid SourceSiteId { get; set; }

        public Guid SourceWebId { get; set; }

        public string SourceWebUrl { get; set; }

        public Guid SourceListId { get; set; }

        public string Title { get; set; }

        public string Description { get; set; }

        public Guid TemplateFeatureId { get; set; }

        public int BaseTemplate { get; set; }

        public string BaseType { get; set; }

        public string RootFolderServerRelativeUrl { get; set; }

        public bool Hidden { get; set; }

        public bool ContentTypesEnabled { get; set; }

        public bool EnableAttachments { get; set; }

        public bool EnableFolderCreation { get; set; }

        public bool EnableVersioning { get; set; }

        public bool EnableMinorVersions { get; set; }

        public bool EnableModeration { get; set; }

        public bool ForceCheckout { get; set; }

        [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
        public ListInformationRightsManagementSnapshot InformationRightsManagement { get; set; }

        public int SourceItemCount { get; set; }

        public IList<ListFieldSnapshot> Fields { get; set; } = new List<ListFieldSnapshot>();

        public IList<ListContentTypeSnapshot> ContentTypes { get; set; } = new List<ListContentTypeSnapshot>();

        public bool HasExplicitUniqueContentTypeOrder { get; set; }

        public IList<string> UniqueContentTypeOrder { get; set; } = new List<string>();

        public IList<ContentTypeSchemaSnapshot> SiteContentTypes { get; set; } = new List<ContentTypeSchemaSnapshot>();

        public IList<ListViewSnapshot> Views { get; set; } = new List<ListViewSnapshot>();

        [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
        public IList<ListViewRenderingResourceSnapshot> ViewRenderingResources { get; set; } = new List<ListViewRenderingResourceSnapshot>();

        public IList<ListItemSnapshot> Items { get; set; } = new List<ListItemSnapshot>();

        public EvidenceAvailability Availability { get; set; } = EvidenceAvailability.Captured;

        public IList<string> Diagnostics { get; set; } = new List<string>();
    }
}
