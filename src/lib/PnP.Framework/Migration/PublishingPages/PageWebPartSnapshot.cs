using System;

namespace PnP.Framework.Migration.PublishingPages
{
    public sealed class PageWebPartSnapshot
    {
        public Guid Id { get; set; }

        public string Title { get; set; }

        public string ZoneId { get; set; }

        public int ZoneIndex { get; set; }

        public bool Hidden { get; set; }

        public string ExportXml { get; set; }

        public string ExportSha256 { get; set; }
    }
}
