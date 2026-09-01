using PnP.Framework.Migration.Evidence;
using PnP.Framework.Migration.Lists.ContentTypes;
using PnP.Framework.Migration.Lists.Fields;
using PnP.Framework.Migration.Lists.Items;
using PnP.Framework.Migration.Lists.Views;
using System;
using System.Collections.Generic;

namespace PnP.Framework.Migration.Lists.Capture
{
    public sealed class ListDependencySnapshot
    {
        public string SchemaVersion { get; set; } = "pnp-list-dependency/v1";

        public Guid SourceSiteId { get; set; }

        public Guid SourceWebId { get; set; }

        public string SourceWebUrl { get; set; }

        public Guid SourceListId { get; set; }

        public string Title { get; set; }

        public string Description { get; set; }

        public Guid TemplateFeatureId { get; set; }

        public int BaseTemplate { get; set; }

        public string BaseType { get; set; }

        public string RootFolderServerRelativeUrl { get; set; }

        public bool Hidden { get; set; }

        public bool ContentTypesEnabled { get; set; }

        public bool EnableAttachments { get; set; }

        public bool EnableFolderCreation { get; set; }

        public bool EnableVersioning { get; set; }

        public bool EnableMinorVersions { get; set; }

        public bool EnableModeration { get; set; }

        public bool ForceCheckout { get; set; }

        public int SourceItemCount { get; set; }

        public IList<ListFieldSnapshot> Fields { get; set; } = new List<ListFieldSnapshot>();

        public IList<ListContentTypeSnapshot> ContentTypes { get; set; } = new List<ListContentTypeSnapshot>();

        public IList<ListViewSnapshot> Views { get; set; } = new List<ListViewSnapshot>();

        public IList<ListItemSnapshot> Items { get; set; } = new List<ListItemSnapshot>();

        public EvidenceAvailability Availability { get; set; } = EvidenceAvailability.Captured;

        public IList<string> Diagnostics { get; set; } = new List<string>();
    }
}
