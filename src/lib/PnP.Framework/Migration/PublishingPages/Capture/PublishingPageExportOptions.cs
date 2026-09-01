namespace PnP.Framework.Migration.PublishingPages.Capture
{
    public sealed class PublishingPageExportOptions
    {
        public string SourcePageServerRelativeUrl { get; set; }

        public bool IncludeWebParts { get; set; } = true;

        public long MaximumDependencyBytes { get; set; } = 10 * 1024 * 1024;
    }
}
