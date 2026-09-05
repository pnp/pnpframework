using System;

namespace PnP.Framework.Migration.Pages.Capture
{
    public sealed class SourcePageFence
    {
        public Guid FileUniqueId { get; set; }

        public string VersionLabel { get; set; }

        public long Length { get; set; }

        public DateTime ModifiedUtc { get; set; }
    }
}
