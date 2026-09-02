using PnP.Framework.Migration.Pages.Publishing.Capture;
using PnP.Framework.Migration.Pages.Publishing.Planning;
using PnP.Framework.Migration.Pages.Publishing.Reporting;
using System;

namespace PnP.Framework.Migration.Pages.Publishing.Packaging
{
    public sealed class PublishingPageMigrationPackage
    {
        public string SchemaVersion { get; set; } = PublishingPagePackageContract.MigrationSchemaVersion;

        public DateTimeOffset PlannedAtUtc { get; set; }

        public string ExportSchemaVersion { get; set; } = PublishingPagePackageContract.ExportSchemaVersion;

        public DateTimeOffset ExportedAtUtc { get; set; }

        public PublishingPagePackageState State { get; set; }

        public PublishingPageWorkflowSelection Selection { get; set; }

        public string SelectionDigest { get; set; }

        public PublishingPageCaptureBundle Snapshot { get; set; }

        public PublishingPageMigrationPlan Plan { get; set; }

        public string SnapshotDigest { get; set; }

        public string PlanDigest { get; set; }

        public PublishingPageMigrationReport Report { get; set; }
    }
}
