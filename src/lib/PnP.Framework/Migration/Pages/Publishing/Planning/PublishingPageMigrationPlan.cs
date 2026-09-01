using PnP.Framework.Migration.Pages.Content;
using PnP.Framework.Migration.Pages.Fields;
using PnP.Framework.Migration.Pages.Lifecycle;
using PnP.Framework.Migration.Pages.Planning;
using PnP.Framework.Migration.Pages.Publishing.Lifecycle;
using PnP.Framework.Migration.Pages.References;
using PnP.Framework.Migration.Pages.Publishing.Verification;
using PnP.Framework.Migration.Verification;
using System.Collections.Generic;

namespace PnP.Framework.Migration.Pages.Publishing.Planning
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

        public PageMigrationOperation Operation { get; set; } = PageMigrationOperation.CreatePage;

        public PublishingPageTargetLifecycle TargetLifecycle { get; set; }

        public string LifecycleReason { get; set; }

        public bool CreateOnly { get; set; } = true;

        public PagePlanningOptions PlanningPolicy { get; set; }

        public PublishingPageTargetSnapshot TargetProbe { get; set; }

        public IList<PageFieldAction> FieldActions { get; set; } = new List<PageFieldAction>();

        public IList<PageReferenceAction> DependencyActions { get; set; } = new List<PageReferenceAction>();

        public IList<PageTextReplacement> Replacements { get; set; } = new List<PageTextReplacement>();

        public string ExpectedPublishingPageContentSha256 { get; set; }

        public IList<string> StorageAssertions { get; set; } = new List<string>();

        public RuntimeVerificationManifest RuntimeVerification { get; set; } = new RuntimeVerificationManifest();

        public IList<string> Blockers { get; set; } = new List<string>();

        public IList<string> Warnings { get; set; } = new List<string>();

        public bool IsExecutable => Blockers.Count == 0;
    }
}
