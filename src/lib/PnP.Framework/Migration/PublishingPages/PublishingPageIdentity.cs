using System;

namespace PnP.Framework.Migration.PublishingPages
{
    public sealed class PublishingPageIdentity
    {
        public string WebUrl { get; set; }

        public string WebServerRelativeUrl { get; set; }

        public string PageServerRelativeUrl { get; set; }

        public int ListItemId { get; set; }

        public Guid FileUniqueId { get; set; }

        public string ContentTypeId { get; set; }

        public string ContentTypeName { get; set; }

        public string VersionLabel { get; set; }

        public long Length { get; set; }

        public DateTime ModifiedUtc { get; set; }

        public string Title { get; set; }

        public string PageLayoutUrl { get; set; }

        public string PageLayoutDescription { get; set; }
    }
}
