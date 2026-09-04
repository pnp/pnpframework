using System.Collections.Generic;

namespace PnP.Framework.Migration.Pages.Publishing.Verification
{
    public sealed class PublishingPageTargetSnapshot
    {
        public string WebUrl { get; set; }

        public string WebServerRelativeUrl { get; set; }

        public string WebTemplate { get; set; }

        public int WebConfiguration { get; set; }

        public string PagesLibraryServerRelativeUrl { get; set; }

        public int PagesLibraryBaseTemplate { get; set; }

        public bool EnableVersioning { get; set; }

        public bool EnableMinorVersions { get; set; }

        public bool EnableModeration { get; set; }

        public bool ForceCheckout { get; set; }

        public string DraftVersionVisibility { get; set; }

        public string PageContentTypeId { get; set; }

        public string PageLayoutUrl { get; set; }

        public bool PageLayoutExists { get; set; }

        public string PreferredTargetPageServerRelativeUrl { get; set; }

        public string TargetPageServerRelativeUrl { get; set; }

        public bool PreferredTargetPageExists { get; set; }

        public bool TargetPathCollisionResolved { get; set; }

        public string TargetPathResolutionReason { get; set; }

        public bool TargetPageExists { get; set; }

        public IList<string> ExistingDependencyPaths { get; set; } = new List<string>();
    }
}
