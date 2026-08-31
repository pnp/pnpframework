namespace PnP.Framework.Migration.PublishingPages
{
    public sealed class PageFieldImportResult
    {
        public string InternalName { get; set; }

        public PageFieldDisposition PlannedDisposition { get; set; }

        public bool Attempted { get; set; }

        public bool Succeeded { get; set; }

        public string Message { get; set; }
    }
}
