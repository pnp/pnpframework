using PnP.Framework.Migration.Pages.Cohorts;
using System;
using System.Collections.Generic;

namespace PnP.Framework.Migration.Pages.Publishing.Profiles
{
    public static class EnterpriseWikiV1CohortPolicy
    {
        public const string CohortId = "enterprise-wiki-v1";

        public const string PolicyVersion = "1";

        public static ValidationCohortAssessment Assess(string contentTypeId)
        {
            if (string.IsNullOrWhiteSpace(contentTypeId))
            {
                return Result(
                    ValidationCohortDisposition.Unknown,
                    "The source ContentTypeId is unavailable, so EW-v1 cohort membership cannot be established.");
            }

            if (contentTypeId.StartsWith(BuiltInContentTypeId.ProjectPage, StringComparison.OrdinalIgnoreCase))
            {
                return Result(
                    ValidationCohortDisposition.Excluded,
                    "Project Page Content Type lineage is intentionally outside the EW-v1 validation cohort; migration capability is assessed independently.");
            }

            if (contentTypeId.StartsWith(BuiltInContentTypeId.EnterpriseWikiPage, StringComparison.OrdinalIgnoreCase))
            {
                return Result(
                    ValidationCohortDisposition.Included,
                    "Enterprise Wiki Content Type lineage is included by the EW-v1 validation policy.");
            }

            return Result(
                ValidationCohortDisposition.Excluded,
                "The source Content Type is outside the EW-v1 validation cohort.");
        }

        public static bool IsIncludedContentType(string contentTypeId)
        {
            return Assess(contentTypeId).Disposition == ValidationCohortDisposition.Included;
        }

        private static ValidationCohortAssessment Result(ValidationCohortDisposition disposition, string reason)
        {
            return new ValidationCohortAssessment
            {
                CohortId = CohortId,
                PolicyVersion = PolicyVersion,
                Disposition = disposition,
                Reasons = new List<string> { reason }
            };
        }
    }
}
