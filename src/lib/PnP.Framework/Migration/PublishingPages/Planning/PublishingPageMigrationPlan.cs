using PnP.Framework.Migration.PublishingPages.Content;
using PnP.Framework.Migration.PublishingPages.Fields;
using PnP.Framework.Migration.PublishingPages.Lifecycle;
using PnP.Framework.Migration.PublishingPages.References;
using PnP.Framework.Migration.PublishingPages.Verification;
using System.Collections.Generic;

namespace PnP.Framework.Migration.PublishingPages.Planning
{
    public sealed class PublishingPageMigrationPlan
    {
        public string SourceSnapshotDigest { get; set; }

        public string SourceWebUrl { get; set; }

        public string SourcePageServerRelativeUrl { get; set; }

        public string TargetWebUrl { get; set; }

        public string TargetWebServerRelativeUrl { get; set; }

        public string TargetPageServerRelativeUrl { get; set; }

        public string PageLayoutName { get; set; }

        public PublishingPageMigrationOperation Operation { get; set; } = PublishingPageMigrationOperation.CreatePage;

        public PublishingPageTargetLifecycle TargetLifecycle { get; set; }

        public string LifecycleReason { get; set; }

        public bool CreateOnly { get; set; } = true;

        public PublishingPagePlanningOptions PlanningPolicy { get; set; }

        public PublishingPageTargetSnapshot TargetProbe { get; set; }

        public IList<PageFieldAction> FieldActions { get; set; } = new List<PageFieldAction>();

        public IList<PageReferenceAction> DependencyActions { get; set; } = new List<PageReferenceAction>();

        public IList<PageTextReplacement> Replacements { get; set; } = new List<PageTextReplacement>();

        public string ExpectedPublishingPageContentSha256 { get; set; }

        public IList<string> StorageAssertions { get; set; } = new List<string>();

        public IList<string> BrowserAssertions { get; set; } = new List<string>();

        public IList<string> Blockers { get; set; } = new List<string>();

        public IList<string> Warnings { get; set; } = new List<string>();

        public bool IsExecutable => Blockers.Count == 0;
    }
}
