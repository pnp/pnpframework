using PnP.Framework.Migration.Pages.Content;
using PnP.Framework.Migration.Pages.Fields;
using PnP.Framework.Migration.Pages.Lifecycle;
using PnP.Framework.Migration.Pages.Planning;
using PnP.Framework.Migration.Pages.Publishing.Lifecycle;
using PnP.Framework.Migration.Pages.Publishing.Layouts;
using PnP.Framework.Migration.Pages.References;
using PnP.Framework.Migration.Pages.Publishing.Verification;
using PnP.Framework.Migration.Verification;
using PnP.Framework.Migration.Topology;
using PnP.Framework.Migration.Lists.Planning;
using PnP.Framework.Migration.Pages.ClassicWebParts.Bindings;
using PnP.Framework.Migration.Pages.Ingredients;
using PnP.Framework.Migration.Diagnostics;
using PnP.Framework.Migration.Taxonomy;
using System.Collections.Generic;
using System.Linq;

namespace PnP.Framework.Migration.Pages.Publishing.Planning
{
    public sealed class PublishingPageMigrationPlan
    {
        public string SourceSnapshotDigest { get; set; }

        public string SourceWebUrl { get; set; }

        public string SourcePageServerRelativeUrl { get; set; }

        public string OriginalIdentifier { get; set; }

        public string TargetWebUrl { get; set; }

        public string TargetWebServerRelativeUrl { get; set; }

        public string PreferredTargetPageServerRelativeUrl { get; set; }

        public string TargetPageServerRelativeUrl { get; set; }

        public bool TargetPathCollisionResolved { get; set; }

        public string TargetPathResolutionReason { get; set; }

        public string PageLayoutName { get; set; }

        public PageMigrationOperation Operation { get; set; } = PageMigrationOperation.CreatePage;

        public PublishingPageTargetLifecycle TargetLifecycle { get; set; }

        public string LifecycleReason { get; set; }

        public bool CreateOnly { get; set; } = true;

        public PagePlanningOptions PlanningPolicy { get; set; }

        public PublishingPageTargetSnapshot TargetProbe { get; set; }

        public PublishingPageLayoutMaterializationPlan LayoutMaterialization { get; set; }

        public PublishingPageLayoutTargetProbe LayoutTargetProbe { get; set; }

        public PublishingPageLayoutTargetAdmission LayoutAdmission { get; set; }

        public IList<PageFieldAction> FieldActions { get; set; } = new List<PageFieldAction>();

        public IList<TaxonomyRelationshipAction> TaxonomyRelationshipActions { get; set; } = new List<TaxonomyRelationshipAction>();

        public IList<PageReferenceAction> DependencyActions { get; set; } = new List<PageReferenceAction>();

        public TopologyPlan Topology { get; set; }

        public TopologyTargetAnalysis TopologyTargetAnalysis { get; set; }

        public ListMigrationPlanSet ListMigration { get; set; }

        public IList<ClassicWebPartAction> WebPartActions { get; set; } = new List<ClassicWebPartAction>();

        public IList<PageTextReplacement> Replacements { get; set; } = new List<PageTextReplacement>();

        public string ExpectedPublishingPageContentSha256 { get; set; }

        public IList<string> StorageAssertions { get; set; } = new List<string>();

        public RuntimeVerificationManifest RuntimeVerification { get; set; } = new RuntimeVerificationManifest();

        public CanonicalPageIngredientGraph IngredientGraph { get; set; }

        public IList<PageIngredientAction> IngredientActions { get; set; } = new List<PageIngredientAction>();

        public PageMigrationOutcome MigrationOutcome { get; set; }

        public IList<MigrationIssue> IngredientIssues { get; set; } = new List<MigrationIssue>();

        public PageIngredientExecutionFrontier ExecutionFrontier { get; set; } = new PageIngredientExecutionFrontier();

        public IList<string> Blockers { get; set; } = new List<string>();

        public IList<string> Warnings { get; set; } = new List<string>();

        public bool IsExecutable => (MigrationOutcome == PageMigrationOutcome.Exact
                || MigrationOutcome == PageMigrationOutcome.ExecutableWithTransform
                || MigrationOutcome == PageMigrationOutcome.ExecutableWithLoss
                || MigrationOutcome == PageMigrationOutcome.PartiallyExecutable)
            && IngredientIssues.All(value => value.Severity != MigrationIssueSeverity.Error)
            && (ExecutionFrontier?.HasExecutableIngredients == true
                || MigrationOutcome == PageMigrationOutcome.ExecutableWithLoss);
    }
}
