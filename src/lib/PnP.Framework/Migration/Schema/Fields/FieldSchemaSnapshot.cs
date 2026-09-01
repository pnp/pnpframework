using PnP.Framework.Migration.Evidence;
using System;
using System.Collections.Generic;

namespace PnP.Framework.Migration.Schema.Fields
{
    public sealed class FieldSchemaSnapshot
    {
        public Guid Id { get; set; }

        public string InternalName { get; set; }

        public string Title { get; set; }

        public string TypeAsString { get; set; }

        public string Group { get; set; }

        public bool Required { get; set; }

        public bool Hidden { get; set; }

        public bool ReadOnly { get; set; }

        public bool Sealed { get; set; }

        public string SchemaXml { get; set; }

        public string SchemaXmlSha256 { get; set; }

        public string PortableSchemaSha256 { get; set; }

        public FieldSchemaRole Role { get; set; }

        public FieldOwnership Ownership { get; set; }

        public TaxonomyFieldBindingSnapshot Taxonomy { get; set; }

        public IList<EvidenceSource> Sources { get; set; } = new List<EvidenceSource>();

        public IList<string> Diagnostics { get; set; } = new List<string>();
    }
}
