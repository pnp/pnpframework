using System;

namespace PnP.Framework.Migration.Pages.ClassicWebParts
{
    public sealed class ClassicWebPartSnapshot
    {
        public Guid Id { get; set; }

        public string Title { get; set; }

        public string TypeName { get; set; }

        public string ZoneId { get; set; }

        public int ZoneIndex { get; set; }

        public bool Hidden { get; set; }

        public string ExportXml { get; set; }

        public string ExportSha256 { get; set; }
    }
}
