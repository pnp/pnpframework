using PnP.Framework.Migration.Evidence;
using PnP.Framework.Migration.Packaging;
using PnP.Framework.Migration.Schema.ContentTypes;
using PnP.Framework.Migration.Pages.Markup;
using System;
using System.Collections.Generic;
using System.Text.Json.Serialization;

namespace PnP.Framework.Migration.Pages.Publishing.Layouts
{
    public sealed class PublishingPageLayoutSnapshot
    {
        public string SchemaVersion { get; set; } = "pnp-publishing-page-layout/v1";

        public PublishingPageLayoutEvidenceState EvidenceState { get; set; }

        public string Url { get; set; }

        public string ServerRelativeUrl { get; set; }

        [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
        public string OwnerSiteCollectionUrl { get; set; }

        [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingDefault)]
        public bool ExternalToPageSiteCollection { get; set; }

        [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
        public LiteralHttpAuthorizationEvidence AuthorizationEvidence { get; set; }

        public string Description { get; set; }

        public Guid FileUniqueId { get; set; }

        public int? CustomizedPageStatus { get; set; }

        public string SetupPath { get; set; }

        public string FileName { get; set; }

        public string ItemContentTypeId { get; set; }

        public string AssociatedContentTypeName { get; set; }

        public string AssociatedContentTypeId { get; set; }

        public string Title { get; set; }

        public string Level { get; set; }

        public string CheckOutType { get; set; }

        public string VersionLabel { get; set; }

        public ArtifactReference Bytes { get; set; }

        public string ContentBase64 { get; set; }

        public PageDirectiveSnapshot PageDirective { get; set; }

        public IList<PublishingPageLayoutRegistration> Registrations { get; set; } = new List<PublishingPageLayoutRegistration>();

        public IList<PublishingPageLayoutControl> Controls { get; set; } = new List<PublishingPageLayoutControl>();

        public IList<PublishingPageLayoutZone> Zones { get; set; } = new List<PublishingPageLayoutZone>();

        public IList<PublishingPageLayoutResourceReference> ResourceReferences { get; set; } = new List<PublishingPageLayoutResourceReference>();

        public IList<PublishingPageLayoutResourceSnapshot> ResourceArtifacts { get; set; } = new List<PublishingPageLayoutResourceSnapshot>();

        public ContentTypeSchemaSnapshot AssociatedContentTypeSchema { get; set; }

        public EvidenceAvailability Availability { get; set; }

        public IList<string> Diagnostics { get; set; } = new List<string>();
    }
}
