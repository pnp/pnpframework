using System;

namespace PnP.Framework.Migration.Pages.Publishing.Verification
{
    public sealed class PublishingPageWebPartVerificationResult
    {
        public Guid SourceWebPartId { get; set; }

        public Guid? TargetWebPartId { get; set; }

        public string ExpectedZoneId { get; set; }

        public int ExpectedZoneIndex { get; set; }

        public string ActualZoneId { get; set; }

        public int? ActualZoneIndex { get; set; }

        public string ExpectedExportSha256 { get; set; }

        public string ActualExportSha256 { get; set; }

        public bool Passed { get; set; }

        public string Message { get; set; }
    }
}
