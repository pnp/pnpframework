namespace PnP.Framework.Migration.Pages.Capture
{
    public sealed class PageCaptureOptions
    {
        public string SourcePageServerRelativeUrl { get; set; }

        public bool IncludeWebParts { get; set; } = true;

        public long MaximumDependencyBytes { get; set; } = 10 * 1024 * 1024;
    }
}
