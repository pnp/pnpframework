using System;

namespace PnP.Framework.Migration.Schema.Fields
{
    public sealed class FieldSchemaTargetProbe
    {
        public Guid FieldId { get; set; }

        public bool Exists { get; set; }

        public string InternalName { get; set; }

        public string Title { get; set; }

        public string TypeAsString { get; set; }

        public string PortableSchemaSha256 { get; set; }

        public bool? UnresolvedTargetTermSetExists { get; set; }

        public string UnresolvedTargetTermSetName { get; set; }
    }
}
