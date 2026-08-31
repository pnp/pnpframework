using System;
using System.Collections.Generic;

namespace PnP.Framework.Migration.PublishingPages.EnterpriseWiki
{
    internal static class EnterpriseWikiMigrationProfile
    {
        public const string SourceProfile = "EnterpriseWiki";

        public const string PageLayoutName = "EnterpriseWiki";

        public const string PageLayoutFileName = "EnterpriseWiki.aspx";

        public const int PublishingPagesListTemplate = 850;

        public static readonly HashSet<string> AdditionalFieldNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            "ArticleByLine",
            "PublishingContact",
            "PublishingPageDescription",
            "PublishingPageImage",
            "PublishingRollupImage",
            "SeoBrowserTitle",
            "SeoKeywords",
            "SeoMetaDescription",
            "Wiki_x0020_Page_x0020_Categories"
        };

        public static readonly HashSet<string> HandledFieldNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            "ContentTypeId",
            "PublishingPageContent",
            "PublishingPageLayout",
            "Title"
        };

        public static bool IsContentType(string contentTypeId)
        {
            if (string.IsNullOrWhiteSpace(contentTypeId))
            {
                return false;
            }

            return contentTypeId.StartsWith(BuiltInContentTypeId.EnterpriseWikiPage, StringComparison.OrdinalIgnoreCase)
                && !contentTypeId.StartsWith(BuiltInContentTypeId.ProjectPage, StringComparison.OrdinalIgnoreCase);
        }
    }
}
