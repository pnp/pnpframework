using Microsoft.SharePoint.Client;
using PnP.Framework.Migration.Packaging;
using PnP.Framework.Migration.Pages.Assessment;
using PnP.Framework.Migration.Pages.Planning;
using PnP.Framework.Migration.Pages.Publishing.Assessment;
using PnP.Framework.Migration.Pages.Publishing.Packaging;
using PnP.Framework.Migration.Pages.Publishing.Planning;
using PnP.Framework.Migration.Pages.Publishing.Profiles;
using PnP.Framework.Migration.Topology;

namespace PnP.Framework.Migration.Pages.Publishing.EnterpriseWiki
{
    public sealed class EnterpriseWikiMigrationPlanner
    {
        private readonly PublishingPageMigrationPlanner planner = new PublishingPageMigrationPlanner();

        private readonly PublishingPageMigrationAssessmentPlanner assessmentPlanner =
            new PublishingPageMigrationAssessmentPlanner();

        public PublishingPageMigrationAssessment Assess(
            PublishingPageExportPackage exportPackage,
            TopologyPlan topology,
            PagePlanningOptions options)
        {
            return assessmentPlanner.Assess(
                exportPackage,
                topology,
                options,
                EnterpriseWikiV1WorkflowPolicy.Instance,
                null);
        }

        public PublishingPageMigrationAssessment Assess(
            PublishingPageExportPackage exportPackage,
            TopologyPlan topology,
            PagePlanningOptions options,
            IMigrationArtifactStore artifactStore)
        {
            return assessmentPlanner.Assess(
                exportPackage,
                topology,
                options,
                EnterpriseWikiV1WorkflowPolicy.Instance,
                artifactStore);
        }

        public PublishingPageMigrationAssessment Assess(
            PublishingPageExportPackage exportPackage,
            TopologyPlan topology,
            PagePlanningOptions options,
            PageAssessmentEvidence assessmentEvidence,
            IMigrationArtifactStore artifactStore = null)
        {
            return assessmentPlanner.Assess(
                exportPackage,
                topology,
                options,
                EnterpriseWikiV1WorkflowPolicy.Instance,
                artifactStore,
                assessmentEvidence);
        }

        public PublishingPageMigrationPackage Plan(
            ClientContext targetContext,
            PublishingPageExportPackage exportPackage,
            PagePlanningOptions options)
        {
            return planner.Plan(targetContext, exportPackage, options, EnterpriseWikiV1WorkflowPolicy.Instance);
        }

        public PublishingPageMigrationPackage Plan(
            ClientContext targetContext,
            PublishingPageExportPackage exportPackage,
            PagePlanningOptions options,
            IMigrationArtifactStore artifactStore)
        {
            return planner.Plan(targetContext, exportPackage, options, EnterpriseWikiV1WorkflowPolicy.Instance, artifactStore);
        }
    }
}
