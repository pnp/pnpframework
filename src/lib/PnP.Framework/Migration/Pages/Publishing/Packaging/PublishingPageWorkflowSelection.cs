using PnP.Framework.Migration.Pages.Cohorts;

namespace PnP.Framework.Migration.Pages.Publishing.Packaging
{
    public sealed class PublishingPageWorkflowSelection
    {
        public string WorkflowId { get; set; }

        public ValidationCohortAssessment ValidationCohort { get; set; }
    }
}
