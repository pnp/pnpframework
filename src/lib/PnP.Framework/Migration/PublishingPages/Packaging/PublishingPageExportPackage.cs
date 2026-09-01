using PnP.Framework.Migration.PublishingPages.Capture;
using System;

namespace PnP.Framework.Migration.PublishingPages.Packaging
{
    public sealed class PublishingPageExportPackage
    {
        public string SchemaVersion { get; set; } = PublishingPagePackageContract.ExportSchemaVersion;

        public DateTimeOffset ExportedAtUtc { get; set; }

        public PublishingPageCaptureBundle Snapshot { get; set; }

        public string SnapshotDigest { get; set; }
    }
}
