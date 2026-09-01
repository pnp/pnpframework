using PnP.Framework.Migration.Packaging;
using PnP.Framework.Migration.Schema.ContentTypes;
using System.Collections.Generic;

namespace PnP.Framework.Migration.Pages.Publishing.Layouts
{
    public sealed class PublishingPageLayoutMaterializationPlan
    {
        public PublishingPageLayoutMaterializationDisposition Disposition { get; set; }

        public string SourceUrl { get; set; }

        public string SourceServerRelativeUrl { get; set; }

        public string SourceFileName { get; set; }

        public ArtifactReference SourceBytes { get; set; }

        public string AssociatedContentTypeName { get; set; }

        public string AssociatedContentTypeId { get; set; }

        public string TargetFileName { get; set; }

        public string TargetPageLayoutName { get; set; }

        public string TargetServerRelativeUrl { get; set; }

        public IList<string> RequiredFieldBindings { get; set; } = new List<string>();

        public IList<PublishingPageLayoutRegistration> RequiredRegistrations { get; set; } = new List<PublishingPageLayoutRegistration>();

        public IList<PublishingPageLayoutZone> Zones { get; set; } = new List<PublishingPageLayoutZone>();

        public IList<PublishingPageLayoutResourceReference> ResourceReferences { get; set; } = new List<PublishingPageLayoutResourceReference>();

        public ContentTypeMaterializationPlan ContentTypeSchema { get; set; }

        public ArtifactReference TargetBytes { get; set; }

        public IList<PublishingPageLayoutResourceMaterializationPlan> ResourceMaterializations { get; set; } = new List<PublishingPageLayoutResourceMaterializationPlan>();

        public IList<PublishingPageLayoutResourceRewrite> ResourceRewrites { get; set; } = new List<PublishingPageLayoutResourceRewrite>();

        public string Reason { get; set; }
    }
}
