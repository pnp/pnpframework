using PnP.Framework.Migration.Evidence;
using System.Collections.Generic;

namespace PnP.Framework.Migration.Lists.Items
{
    public enum ListItemValueKind
    {
        Null = 0,
        String = 1,
        StringCollection = 2,
        Boolean = 3,
        Number = 4,
        DateTime = 5,
        Guid = 6,
        Url = 7,
        User = 8,
        UserCollection = 9,
        Lookup = 10,
        LookupCollection = 11,
        Taxonomy = 12,
        TaxonomyCollection = 13,
        ByteArray = 14,
        Unsupported = 15
    }

    public sealed class ListItemUrlValueSnapshot
    {
        public string Url { get; set; }

        public string Description { get; set; }
    }

    public sealed class ListItemLookupValueSnapshot
    {
        public int LookupId { get; set; }

        public string LookupValue { get; set; }
    }

    public sealed class ListItemTaxonomyValueSnapshot
    {
        public string Label { get; set; }

        public string TermGuid { get; set; }

        public int WssId { get; set; }
    }

    public sealed class ListItemValueSnapshot
    {
        public string InternalName { get; set; }

        public ListItemValueKind Kind { get; set; }

        public string ScalarValue { get; set; }

        public IList<string> StringValues { get; set; } = new List<string>();

        public ListItemUrlValueSnapshot UrlValue { get; set; }

        public IList<ListItemLookupValueSnapshot> LookupValues { get; set; } = new List<ListItemLookupValueSnapshot>();

        public IList<ListItemTaxonomyValueSnapshot> TaxonomyValues { get; set; } = new List<ListItemTaxonomyValueSnapshot>();

        public string RawType { get; set; }

        public string RawValue { get; set; }

        public string RawValueJson { get; set; }

        public EvidenceAvailability Availability { get; set; } = EvidenceAvailability.Captured;

        public IList<string> Diagnostics { get; set; } = new List<string>();
    }
}
