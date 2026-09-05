using System;

namespace PnP.Framework.Migration.Pages.Publishing.Layouts
{
    internal sealed class PublishingPageLayoutSourceLocation
    {
        public Uri AbsoluteLayoutUrl { get; private set; }

        public Uri OwnerSiteCollectionUrl { get; private set; }

        public string ServerRelativeUrl { get; private set; }

        public bool IsExternalToPageSiteCollection { get; private set; }

        public static bool TryResolve(
            Uri pageSiteCollectionUrl,
            string authoredLayoutUrl,
            out PublishingPageLayoutSourceLocation location,
            out string diagnostic)
        {
            location = null;
            diagnostic = null;
            if (pageSiteCollectionUrl == null || !pageSiteCollectionUrl.IsAbsoluteUri)
            {
                diagnostic = "An absolute source page Site Collection URL is required.";
                return false;
            }

            if (string.IsNullOrWhiteSpace(authoredLayoutUrl))
            {
                diagnostic = "PublishingPageLayout is unavailable on the source page.";
                return false;
            }

            Uri absoluteLayoutUrl;
            if (!Uri.TryCreate(authoredLayoutUrl.Trim(), UriKind.Absolute, out absoluteLayoutUrl))
            {
                var ownerBase = new Uri(pageSiteCollectionUrl.AbsoluteUri.TrimEnd('/') + "/");
                if (!Uri.TryCreate(ownerBase, authoredLayoutUrl.Trim(), out absoluteLayoutUrl))
                {
                    diagnostic = "PublishingPageLayout could not be resolved to an absolute URL.";
                    return false;
                }
            }

            if (!string.Equals(absoluteLayoutUrl.Scheme, Uri.UriSchemeHttps, StringComparison.OrdinalIgnoreCase)
                || !SameOrigin(pageSiteCollectionUrl, absoluteLayoutUrl))
            {
                diagnostic = "PublishingPageLayout points outside the source tenant HTTPS origin.";
                return false;
            }

            var serverRelativeUrl = Uri.UnescapeDataString(absoluteLayoutUrl.AbsolutePath);
            const string galleryMarker = "/_catalogs/masterpage/";
            var marker = serverRelativeUrl.IndexOf(galleryMarker, StringComparison.OrdinalIgnoreCase);
            if (marker < 0)
            {
                diagnostic = "PublishingPageLayout is not under a Site Collection master page gallery.";
                return false;
            }

            var ownerPath = serverRelativeUrl.Substring(0, marker);
            var ownerUrl = new Uri(
                absoluteLayoutUrl.GetLeftPart(UriPartial.Authority)
                + (string.IsNullOrEmpty(ownerPath) ? "/" : ownerPath));
            location = new PublishingPageLayoutSourceLocation
            {
                AbsoluteLayoutUrl = absoluteLayoutUrl,
                OwnerSiteCollectionUrl = ownerUrl,
                ServerRelativeUrl = serverRelativeUrl,
                IsExternalToPageSiteCollection = !SameAbsoluteUrl(pageSiteCollectionUrl, ownerUrl)
            };
            return true;
        }

        private static bool SameOrigin(Uri left, Uri right)
        {
            return string.Equals(left.Scheme, right.Scheme, StringComparison.OrdinalIgnoreCase)
                && string.Equals(left.Host, right.Host, StringComparison.OrdinalIgnoreCase)
                && left.Port == right.Port;
        }

        private static bool SameAbsoluteUrl(Uri left, Uri right)
        {
            return string.Equals(
                left.AbsoluteUri.TrimEnd('/'),
                right.AbsoluteUri.TrimEnd('/'),
                StringComparison.OrdinalIgnoreCase);
        }
    }
}
