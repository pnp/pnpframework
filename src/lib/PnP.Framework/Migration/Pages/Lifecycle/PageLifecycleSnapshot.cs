using System;

namespace PnP.Framework.Migration.Pages.Lifecycle
{
    public sealed class PageLifecycleSnapshot
    {
        public string CheckOutType { get; set; }

        public string Level { get; set; }

        public int? ModerationStatus { get; set; }

        public DateTime CreatedUtc { get; set; }

        public DateTime ModifiedUtc { get; set; }
    }
}
