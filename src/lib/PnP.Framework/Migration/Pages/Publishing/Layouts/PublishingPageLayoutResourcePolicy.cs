using System;

namespace PnP.Framework.Migration.Pages.Publishing.Layouts
{
    internal static class PublishingPageLayoutResourcePolicy
    {
        public static bool IsTargetRuntimeResource(string value)
        {
            if (string.IsNullOrWhiteSpace(value) || value.StartsWith("#", StringComparison.Ordinal))
            {
                return true;
            }

            var normalized = value.Trim().Replace('\\', '/');
            Uri absolute;
            if (Uri.TryCreate(normalized, UriKind.Absolute, out absolute))
            {
                return absolute.Host.EndsWith(".sharepoint.com", StringComparison.OrdinalIgnoreCase)
                    && (absolute.AbsolutePath.IndexOf("/_layouts/", StringComparison.OrdinalIgnoreCase) >= 0
                        || absolute.AbsolutePath.IndexOf("/_controltemplates/", StringComparison.OrdinalIgnoreCase) >= 0);
            }

            if (normalized.StartsWith("/_layouts/", StringComparison.OrdinalIgnoreCase)
                || normalized.StartsWith("/_controltemplates/", StringComparison.OrdinalIgnoreCase)
                || normalized.StartsWith("~site/_layouts/", StringComparison.OrdinalIgnoreCase)
                || normalized.StartsWith("~site/_controltemplates/", StringComparison.OrdinalIgnoreCase)
                || normalized.StartsWith("~sitecollection/_layouts/", StringComparison.OrdinalIgnoreCase)
                || normalized.StartsWith("~sitecollection/_controltemplates/", StringComparison.OrdinalIgnoreCase))
            {
                return true;
            }

            var path = normalized.Split('?', '#')[0];
            const string prefix = "~sitecollection/Style Library/~language/";
            if (!path.StartsWith(prefix, StringComparison.OrdinalIgnoreCase))
            {
                return false;
            }

            var coreStylesPath = path.Substring(prefix.Length);
            const string themable = "Themable/";
            if (coreStylesPath.StartsWith(themable, StringComparison.OrdinalIgnoreCase))
            {
                coreStylesPath = coreStylesPath.Substring(themable.Length);
            }

            const string coreStyles = "Core Styles/";
            if (!coreStylesPath.StartsWith(coreStyles, StringComparison.OrdinalIgnoreCase))
            {
                return false;
            }

            var fileName = coreStylesPath.Substring(coreStyles.Length);
            return string.Equals(fileName, "editmode15.css", StringComparison.OrdinalIgnoreCase)
                || string.Equals(fileName, "pagelayouts15.css", StringComparison.OrdinalIgnoreCase)
                || string.Equals(fileName, "page-layouts-21.css", StringComparison.OrdinalIgnoreCase)
                || string.Equals(fileName, "edit-mode-21.css", StringComparison.OrdinalIgnoreCase);
        }

        public static Uri ResolveSourceUri(Uri sourceWebUrl, Uri sourceSiteCollectionUrl, string reference)
        {
            if (sourceWebUrl == null)
            {
                throw new ArgumentNullException(nameof(sourceWebUrl));
            }

            if (sourceSiteCollectionUrl == null)
            {
                throw new ArgumentNullException(nameof(sourceSiteCollectionUrl));
            }

            if (IsTargetRuntimeResource(reference))
            {
                return null;
            }

            var value = (reference ?? string.Empty).Trim().Replace('\\', '/');
            Uri absolute;
            if (Uri.TryCreate(value, UriKind.Absolute, out absolute))
            {
                return string.Equals(absolute.Scheme, Uri.UriSchemeHttps, StringComparison.OrdinalIgnoreCase) ? absolute : null;
            }

            var webPath = sourceWebUrl.AbsolutePath.TrimEnd('/');
            var sitePath = sourceSiteCollectionUrl.AbsolutePath.TrimEnd('/');
            string path = null;
            if (value.StartsWith("~sitecollection/", StringComparison.OrdinalIgnoreCase))
            {
                path = sitePath + "/" + value.Substring("~sitecollection/".Length);
            }
            else if (value.StartsWith("~site/", StringComparison.OrdinalIgnoreCase))
            {
                path = webPath + "/" + value.Substring("~site/".Length);
            }
            else if (value.StartsWith("~/", StringComparison.Ordinal))
            {
                path = webPath + "/" + value.Substring(2);
            }
            else if (value.StartsWith("/", StringComparison.Ordinal))
            {
                path = value;
            }

            return path == null ? null : new Uri(new Uri(sourceWebUrl.GetLeftPart(UriPartial.Authority)), path);
        }

        public static bool IsWebOwnedAsset(Uri webUrl, Uri candidate)
        {
            if (webUrl == null || candidate == null
                || !string.Equals(webUrl.Scheme, candidate.Scheme, StringComparison.OrdinalIgnoreCase)
                || !string.Equals(webUrl.Host, candidate.Host, StringComparison.OrdinalIgnoreCase)
                || webUrl.Port != candidate.Port)
            {
                return false;
            }

            var webPath = Uri.UnescapeDataString(webUrl.AbsolutePath).TrimEnd('/');
            var candidatePath = Uri.UnescapeDataString(candidate.AbsolutePath);
            return candidatePath.StartsWith(webPath + "/SiteAssets/", StringComparison.OrdinalIgnoreCase)
                || candidatePath.StartsWith(webPath + "/Style Library/", StringComparison.OrdinalIgnoreCase);
        }
    }
}
