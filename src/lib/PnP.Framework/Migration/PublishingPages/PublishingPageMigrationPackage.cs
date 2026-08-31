using System;

namespace PnP.Framework.Migration.PublishingPages
{
    public sealed class PublishingPageMigrationPackage
    {
        public string SchemaVersion { get; set; } = PublishingPagePackageContract.MigrationSchemaVersion;

        public DateTimeOffset PlannedAtUtc { get; set; }

        public string ExportSchemaVersion { get; set; } = PublishingPagePackageContract.ExportSchemaVersion;

        public DateTimeOffset ExportedAtUtc { get; set; }

        public PublishingPagePackageState State { get; set; }

        public PublishingPageCaptureBundle Snapshot { get; set; }

        public PublishingPageMigrationPlan Plan { get; set; }

        public string SnapshotDigest { get; set; }

        public string PlanDigest { get; set; }

        public PublishingPageMigrationReport Report { get; set; }
    }
}
