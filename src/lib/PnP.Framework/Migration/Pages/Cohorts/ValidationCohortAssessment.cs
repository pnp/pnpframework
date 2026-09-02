using System.Collections.Generic;

namespace PnP.Framework.Migration.Pages.Cohorts
{
    public sealed class ValidationCohortAssessment
    {
        public string CohortId { get; set; }

        public string PolicyVersion { get; set; }

        public ValidationCohortDisposition Disposition { get; set; }

        public IList<string> Reasons { get; set; } = new List<string>();
    }
}
