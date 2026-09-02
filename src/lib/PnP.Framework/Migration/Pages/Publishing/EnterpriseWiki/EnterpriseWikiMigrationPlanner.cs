using Microsoft.SharePoint.Client;
using PnP.Framework.Migration.Packaging;
using PnP.Framework.Migration.Pages.Planning;
using PnP.Framework.Migration.Pages.Publishing.Packaging;
using PnP.Framework.Migration.Pages.Publishing.Planning;
using PnP.Framework.Migration.Pages.Publishing.Profiles;

namespace PnP.Framework.Migration.Pages.Publishing.EnterpriseWiki
{
    public sealed class EnterpriseWikiMigrationPlanner
    {
        private readonly PublishingPageMigrationPlanner planner = new PublishingPageMigrationPlanner();

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
