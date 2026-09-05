using PnP.Framework.Migration.Evidence;
using System.Collections.Generic;

namespace PnP.Framework.Migration.Pages.Publishing.Layouts
{
    public sealed class PublishingPageLayoutResourceTargetProbe
    {
        public string TargetServerRelativeUrl { get; set; }

        public bool FileExists { get; set; }

        public string ExistingBytesSha256 { get; set; }

        public bool CanWrite { get; set; }

        public EvidenceAvailability Availability { get; set; }

        public IList<string> Diagnostics { get; set; } = new List<string>();
    }
}
