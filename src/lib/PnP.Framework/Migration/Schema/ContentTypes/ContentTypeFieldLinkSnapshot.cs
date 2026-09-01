using PnP.Framework.Migration.Schema.Fields;
using System;

namespace PnP.Framework.Migration.Schema.ContentTypes
{
    public sealed class ContentTypeFieldLinkSnapshot
    {
        public Guid FieldId { get; set; }

        public string Name { get; set; }

        public bool Required { get; set; }

        public bool Hidden { get; set; }

        public FieldSchemaRole Role { get; set; }
    }
}
