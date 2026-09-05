using PnP.Framework.Migration.Evidence;
using System;
using System.Collections.Generic;
using System.Text.Json.Serialization;

namespace PnP.Framework.Migration.Lists.Views
{
    public sealed class ListViewSnapshot
    {
        public Guid Id { get; set; }

        public string Title { get; set; }

        public string ServerRelativeUrl { get; set; }

        public bool Hidden { get; set; }

        public bool DefaultView { get; set; }

        public bool PersonalView { get; set; }

        public string ViewType { get; set; }

        public uint RowLimit { get; set; }

        public bool Paged { get; set; }

        public string ViewQuery { get; set; }

        public IList<string> ViewFields { get; set; } = new List<string>();

        public string ListViewXml { get; set; }

        public string ListViewXmlSha256 { get; set; }

        public string JsLink { get; set; }

        public string XslLink { get; set; }

        [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
        public IList<ListViewRenderingResourceBindingSnapshot> RenderingResourceBindings { get; set; } = new List<ListViewRenderingResourceBindingSnapshot>();

        public bool IsPageBound { get; set; }

        public EvidenceAvailability Availability { get; set; } = EvidenceAvailability.Captured;

        public IList<string> Diagnostics { get; set; } = new List<string>();
    }
}
