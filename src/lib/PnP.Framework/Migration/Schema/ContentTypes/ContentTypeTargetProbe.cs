using PnP.Framework.Migration.Evidence;
using PnP.Framework.Migration.Schema.Fields;
using System;
using System.Collections.Generic;

namespace PnP.Framework.Migration.Schema.ContentTypes
{
    public sealed class ContentTypeTargetProbe
    {
        public string ContentTypeId { get; set; }

        public bool ParentContentTypeAvailable { get; set; }

        public string ResolvedParentContentTypeId { get; set; }

        public IList<ContentTypeFieldLinkTargetProbe> ParentFieldLinks { get; set; } = new List<ContentTypeFieldLinkTargetProbe>();

        public bool ContentTypeExists { get; set; }

        public string ExistingName { get; set; }

        public string ExistingDescription { get; set; }

        public string ExistingGroup { get; set; }

        public bool ExistingReadOnly { get; set; }

        public bool ExistingSealed { get; set; }

        public bool ExistingHidden { get; set; }

        public string ExistingParentContentTypeId { get; set; }

        public IList<ContentTypeFieldLinkTargetProbe> ExistingFieldLinks { get; set; } = new List<ContentTypeFieldLinkTargetProbe>();

        public IList<string> SameNameDifferentIds { get; set; } = new List<string>();

        public IList<FieldSchemaTargetProbe> Fields { get; set; } = new List<FieldSchemaTargetProbe>();

        public bool CanManageContentTypes { get; set; }

        public EvidenceAvailability Availability { get; set; }

        public IList<string> Diagnostics { get; set; } = new List<string>();
    }
}
