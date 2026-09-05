using PnP.Framework.Migration.Pages.Assessment;
using PnP.Framework.Migration.Pages.Ingredients;
using PnP.Framework.Migration.Pages.Planning;
using PnP.Framework.Migration.Pages.Publishing.Lifecycle;
using System.Collections.Generic;

namespace PnP.Framework.Migration.Pages.Publishing.Assessment
{
    /// <summary>
    /// A deterministic, source-authoritative planning projection that can be
    /// produced before the target topology exists. It identifies known actions,
    /// target inspections, engineering work, and literal authorization boundaries,
    /// but cannot be imported.
    /// </summary>
    public sealed class PublishingPageMigrationAssessment
    {
        public string SchemaVersion { get; set; } = "pnp-publishing-page-migration-assessment/v2";

        public string SourceSnapshotDigest { get; set; }

        public string WorkflowId { get; set; }

        public string SelectionDigest { get; set; }

        public string SourceWebUrl { get; set; }

        public string SourcePageServerRelativeUrl { get; set; }

        public string TargetSiteCollectionUrl { get; set; }

        public string TargetWebUrl { get; set; }

        public string TargetWebServerRelativeUrl { get; set; }

        public string TargetPageServerRelativeUrl { get; set; }

        public string TopologyPlanDigest { get; set; }

        public PagePlanningOptions PlanningPolicy { get; set; }

        public PublishingPageTargetLifecycle TargetLifecycle { get; set; }

        public string LifecycleReason { get; set; }

        public CanonicalPageIngredientGraph IngredientGraph { get; set; }

        public IList<PageIngredientAssessment> IngredientAssessments { get; set; } = new List<PageIngredientAssessment>();

        public IList<string> KnownGaps { get; set; } = new List<string>();

        public IList<string> Warnings { get; set; } = new List<string>();

        public PageMigrationAssessmentState State { get; set; }

        public string AssessmentDigest { get; set; }
    }
}
