using System;
using System.Collections.Generic;

namespace PnP.Framework.Migration.Lists.ContentTypes
{
    public sealed class ListContentTypeFieldLinkSnapshot
    {
        public Guid FieldId { get; set; }

        public string InternalName { get; set; }

        public string DisplayName { get; set; }

        public bool Required { get; set; }

        public bool Hidden { get; set; }

        public bool ReadOnly { get; set; }
    }

    public sealed class ListContentTypeSnapshot
    {
        public string Id { get; set; }

        public string Name { get; set; }

        public string Description { get; set; }

        public string Group { get; set; }

        public string ParentId { get; set; }

        public bool Hidden { get; set; }

        public bool ReadOnly { get; set; }

        public bool Sealed { get; set; }

        public IList<ListContentTypeFieldLinkSnapshot> FieldLinks { get; set; } = new List<ListContentTypeFieldLinkSnapshot>();
    }
}
