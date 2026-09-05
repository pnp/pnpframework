using System;
using System.Collections.Generic;

namespace PnP.Framework.Migration.Pages.Publishing.Profiles
{
    internal static class EnterpriseWikiV1WorkflowPolicy
    {
        public static readonly PublishingPageWorkflowPolicy Instance = new PublishingPageWorkflowPolicy
        {
            WorkflowId = EnterpriseWikiV1CohortPolicy.CohortId,
            PreferredTargetPageLayoutFileName = "EnterpriseWiki.aspx",
            FieldsHandledByPageWriter = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
            {
                "ContentTypeId",
                "FileLeafRef",
                "PublishingPageContent",
                "PublishingPageLayout",
                "Title"
            },
            RecognizedPageFields = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
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
            },
            AssessValidationCohort = EnterpriseWikiV1CohortPolicy.Assess
        };
    }
}
