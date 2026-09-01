using PnP.Framework.Migration.Schema.Fields;
using System.Collections.Generic;

namespace PnP.Framework.Migration.Schema.ContentTypes
{
    public sealed class ContentTypeMaterializationPlan
    {
        public ContentTypeMaterializationDisposition Disposition { get; set; }

        public string SourceWebUrl { get; set; }

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

        public IList<FieldSchemaMaterializationPlan> Fields { get; set; } = new List<FieldSchemaMaterializationPlan>();

        public string Reason { get; set; }
    }
}
