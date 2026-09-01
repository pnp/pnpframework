using PnP.Framework.Migration.Packaging;

namespace PnP.Framework.Migration.Pages.Publishing.Layouts
{
    public sealed class PublishingPageLayoutResourceMaterializationPlan
    {
        public string SourceReference { get; set; }

        public string SourceUrl { get; set; }

        public PublishingPageLayoutResourceEvidenceState SourceEvidenceState { get; set; }

        public PublishingPageLayoutResourceMaterializationDisposition Disposition { get; set; }

        public ArtifactReference SourceArtifact { get; set; }

        public string SourceContentBase64 { get; set; }

        public string TargetServerRelativeUrl { get; set; }

        public string TargetReference { get; set; }

        public string Reason { get; set; }
    }
}
