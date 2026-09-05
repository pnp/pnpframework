using PnP.Framework.Migration.Pages.Capture;
using PnP.Framework.Migration.Pages.Fields;
using PnP.Framework.Migration.Pages.Lifecycle;
using PnP.Framework.Migration.Pages.References;
using PnP.Framework.Migration.Pages.Security;
using PnP.Framework.Migration.Pages.ClassicWebParts;
using PnP.Framework.Migration.Pages.Publishing.Layouts;
using PnP.Framework.Migration.Pages.ClassicWebParts.Bindings;
using PnP.Framework.Migration.Lists.Capture;
using PnP.Framework.Migration.Lists.Planning;
using PnP.Framework.Migration.Topology;
using PnP.Framework.Migration.Pages.Markup;
using PnP.Framework.Migration.Pages.Runtime;
using PnP.Framework.Migration.Pages.Profiles;
using PnP.Framework.Migration.Pages.Ingredients;
using System.Collections.Generic;

namespace PnP.Framework.Migration.Pages.Publishing.Capture
{
    public sealed class PublishingPageCaptureBundle
    {
        public PageCaptureOptions CapturePolicy { get; set; }

        public PageIdentity Source { get; set; }

        public PageArtifactSnapshot PageArtifact { get; set; }

        public PageRuntimeSnapshot Runtime { get; set; }

        public IList<PageProfileSignal> ProfileSignals { get; set; } = new List<PageProfileSignal>();

        public CanonicalPageIngredientGraph IngredientGraph { get; set; }

        public PublishingPageLayoutSnapshot Layout { get; set; }

        public string PublishingPageContent { get; set; }

        public string PublishingPageContentSha256 { get; set; }

        public IList<PageFieldValueSnapshot> Fields { get; set; } = new List<PageFieldValueSnapshot>();

        public IList<ClassicWebPartSnapshot> WebParts { get; set; } = new List<ClassicWebPartSnapshot>();

        public IList<ClassicListWebPartBindingSnapshot> ListWebPartBindings { get; set; } = new List<ClassicListWebPartBindingSnapshot>();

        public IList<ListDependencySnapshot> ListDependencies { get; set; } = new List<ListDependencySnapshot>();

        public IList<ListLookupDependency> ListLookupDependencies { get; set; } = new List<ListLookupDependency>();

        public SourceSiteCollectionSnapshot SourceTopology { get; set; }

        public IList<PageReferenceSnapshot> Dependencies { get; set; } = new List<PageReferenceSnapshot>();

        public PageSecuritySnapshot Security { get; set; }

        public PageLifecycleSnapshot Lifecycle { get; set; }

        public SourcePageFence SourceFence { get; set; }

        public IList<string> Blockers { get; set; } = new List<string>();

        public IList<string> Warnings { get; set; } = new List<string>();
    }
}
