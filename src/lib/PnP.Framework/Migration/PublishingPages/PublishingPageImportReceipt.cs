using System;
using System.Collections.Generic;

namespace PnP.Framework.Migration.PublishingPages
{
    public sealed class PublishingPageImportReceipt
    {
        public string SchemaVersion { get; set; } = PublishingPagePackageContract.ReceiptSchemaVersion;

        public DateTimeOffset StartedAtUtc { get; set; }

        public DateTimeOffset CompletedAtUtc { get; set; }

        public string ApprovedPlanDigest { get; set; }

        public string TargetWebUrl { get; set; }

        public string TargetPageServerRelativeUrl { get; set; }

        public Guid TargetFileUniqueId { get; set; }

        public int TargetListItemId { get; set; }

        public string TargetContentTypeId { get; set; }

        public string TargetVersionLabel { get; set; }

        public PublishingPageTargetLifecycle ExpectedLifecycle { get; set; }

        public string ActualFileLevel { get; set; }

        public string ActualCheckOutType { get; set; }

        public int? ActualModerationStatus { get; set; }

        public bool LifecycleMatched { get; set; }

        public string ExpectedPublishingPageContentSha256 { get; set; }

        public string PersistedPublishingPageContentSha256 { get; set; }

        public bool StorageContentEqual { get; set; }

        public int ImportedWebPartCount { get; set; }

        public int MaterializedDependencyCount { get; set; }

        public IList<PageFieldImportResult> FieldResults { get; set; } = new List<PageFieldImportResult>();

        public bool FreshReadbackPassed { get; set; }

        public IList<string> Warnings { get; set; } = new List<string>();

        public bool Succeeded { get; set; }
    }
}
