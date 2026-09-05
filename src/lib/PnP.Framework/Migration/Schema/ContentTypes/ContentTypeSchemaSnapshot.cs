using PnP.Framework.Migration.Evidence;
using PnP.Framework.Migration.Schema.Fields;
using System.Collections.Generic;

namespace PnP.Framework.Migration.Schema.ContentTypes
{
    public sealed class ContentTypeSchemaSnapshot
    {
        public string SchemaVersion { get; set; } = "pnp-content-type-schema/v1";

        public ContentTypeSchemaEvidenceState EvidenceState { get; set; }

        public string SourceWebUrl { get; set; }

        public string SourceScope { get; set; }

        public string ContentTypeId { get; set; }

        public string Name { get; set; }

        public string Description { get; set; }

        public string Group { get; set; }

        public bool ReadOnly { get; set; }

        public bool Sealed { get; set; }

        public bool Hidden { get; set; }

        public string ParentContentTypeId { get; set; }

        public string ParentContentTypeName { get; set; }

        public IList<ContentTypeFieldLinkSnapshot> RequiredFieldLinks { get; set; } = new List<ContentTypeFieldLinkSnapshot>();

        public IList<FieldSchemaSnapshot> RequiredFieldClosure { get; set; } = new List<FieldSchemaSnapshot>();

        public EvidenceAvailability Availability { get; set; }

        public IList<EvidenceSource> Sources { get; set; } = new List<EvidenceSource>();

        public IList<string> Diagnostics { get; set; } = new List<string>();
    }
}
