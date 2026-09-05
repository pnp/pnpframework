using PnP.Framework.Migration.Evidence;
using PnP.Framework.Migration.Schema.ContentTypes;
using System.Collections.Generic;

namespace PnP.Framework.Migration.Pages.Publishing.Layouts
{
    public sealed class PublishingPageLayoutTargetProbe
    {
        public string TargetServerRelativeUrl { get; set; }

        public bool FileExists { get; set; }

        public string ExistingBytesSha256 { get; set; }

        public string ExistingAssociatedContentTypeName { get; set; }

        public string ExistingAssociatedContentTypeId { get; set; }

        public bool AssociatedContentTypeAvailable { get; set; }

        public string ResolvedAssociatedContentTypeId { get; set; }

        public IList<string> MissingFieldBindings { get; set; } = new List<string>();

        public bool CanAddAndCustomizePages { get; set; }

        public ContentTypeTargetProbe ContentTypeSchema { get; set; }

        public IList<PublishingPageLayoutResourceTargetProbe> Resources { get; set; } = new List<PublishingPageLayoutResourceTargetProbe>();

        public EvidenceAvailability Availability { get; set; }

        public IList<string> Diagnostics { get; set; } = new List<string>();
    }
}
