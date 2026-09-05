using PnP.Framework.Migration.Evidence;
using PnP.Framework.Migration.Schema.Fields;
using System;
using System.Collections.Generic;

namespace PnP.Framework.Migration.Lists.Fields
{
    public sealed class ListFieldSnapshot
    {
        public Guid Id { get; set; }

        public string InternalName { get; set; }

        public string Title { get; set; }

        public string TypeAsString { get; set; }

        public string Group { get; set; }

        public string SchemaXml { get; set; }

        public string SchemaXmlSha256 { get; set; }

        public string PortableSchemaSha256 { get; set; }

        public bool Hidden { get; set; }

        public bool ReadOnly { get; set; }

        public bool Required { get; set; }

        public bool FromBaseType { get; set; }

        public bool Sealed { get; set; }

        public Guid? SourceLookupWebId { get; set; }

        public Guid? SourceLookupListId { get; set; }

        public string LookupField { get; set; }

        public TaxonomyFieldBindingSnapshot Taxonomy { get; set; }

        public EvidenceAvailability Availability { get; set; } = EvidenceAvailability.Captured;

        public IList<string> Diagnostics { get; set; } = new List<string>();
    }
}
