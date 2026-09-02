using PnP.Framework.Migration.Pages.Publishing.Capture;
using System;

namespace PnP.Framework.Migration.Pages.Publishing.Packaging
{
    public sealed class PublishingPageExportPackage
    {
        public string SchemaVersion { get; set; } = PublishingPagePackageContract.ExportSchemaVersion;

        public DateTimeOffset ExportedAtUtc { get; set; }

        public PublishingPageWorkflowSelection Selection { get; set; }

        public string SelectionDigest { get; set; }

        public PublishingPageCaptureBundle Snapshot { get; set; }

        public string SnapshotDigest { get; set; }
    }
}
