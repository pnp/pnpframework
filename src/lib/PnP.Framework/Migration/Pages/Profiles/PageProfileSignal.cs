namespace PnP.Framework.Migration.Pages.Profiles
{
    public sealed class PageProfileSignal
    {
        public string ProfileId { get; set; }

        public PageProfileSignalKind Kind { get; set; }

        public string Subject { get; set; }

        public string Evidence { get; set; }
    }
}
