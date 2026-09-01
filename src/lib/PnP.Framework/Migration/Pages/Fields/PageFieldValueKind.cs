namespace PnP.Framework.Migration.Pages.Fields
{
    public enum PageFieldValueKind
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
}
