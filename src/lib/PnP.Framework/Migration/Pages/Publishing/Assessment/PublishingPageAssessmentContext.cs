using PnP.Framework.Migration.Lists.Planning;
using PnP.Framework.Migration.Pages.ClassicWebParts.Bindings;
using PnP.Framework.Migration.Pages.Content;
using PnP.Framework.Migration.Pages.Planning;
using PnP.Framework.Migration.Pages.Publishing.Capture;
using PnP.Framework.Migration.Pages.Publishing.Layouts;
using PnP.Framework.Migration.Pages.Publishing.Lifecycle;
using PnP.Framework.Migration.Pages.Publishing.Profiles;
using PnP.Framework.Migration.Pages.References;
using PnP.Framework.Migration.Taxonomy.Assets;
using PnP.Framework.Migration.Topology;
using System.Collections.Generic;

namespace PnP.Framework.Migration.Pages.Publishing.Assessment
{
    internal sealed class PublishingPageAssessmentContext
    {
        public PublishingPageCaptureBundle Snapshot { get; set; }

        public PublishingPageWorkflowPolicy WorkflowPolicy { get; set; }

        public PagePlanningOptions Options { get; set; }

        public TaxonomyAssetReviewPlan TaxonomyAssetReviewPlan { get; set; }

        public SiteCollectionMappingPlan TargetSite { get; set; }

        public WebMappingPlan TargetWeb { get; set; }

        public string TargetPageServerRelativeUrl { get; set; }

        public PublishingPageLayoutMaterializationPlan LayoutPlan { get; set; }

        public string LayoutPlanningFailure { get; set; }

        public PublishingPageTargetLifecycle TargetLifecycle { get; set; }

        public string LifecycleReason { get; set; }

        public IList<PageTextReplacement> Replacements { get; set; } = new List<PageTextReplacement>();

        public IList<PageReferenceAction> ReferenceActions { get; set; } = new List<PageReferenceAction>();

        public string ReferencePlanningFailure { get; set; }

        public ListMigrationPlanSet ListPlan { get; set; }

        public string ListPlanningFailure { get; set; }

        public IList<ClassicWebPartAction> WebPartActions { get; set; } = new List<ClassicWebPartAction>();

        public IList<string> KnownGaps { get; set; } = new List<string>();

        public IList<string> Warnings { get; set; } = new List<string>();
    }
}
