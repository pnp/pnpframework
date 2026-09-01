using System;

namespace PnP.Framework.Migration.Schema.ContentTypes
{
    public sealed class ContentTypeFieldLinkTargetProbe
    {
        public Guid FieldId { get; set; }

        public string Name { get; set; }

        public bool Required { get; set; }

        public bool Hidden { get; set; }
    }
}
