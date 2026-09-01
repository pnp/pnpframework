using PnP.Framework.Migration.Diagnostics;
using System;
using System.Collections.Generic;
using System.Linq;

namespace PnP.Framework.Migration.Pages.ClassicWebParts.Bindings
{
    public sealed class ClassicListWebPartBindingSnapshot
    {
        public Guid SourceWebPartId { get; set; }

        public string TypeName { get; set; }

        public string Title { get; set; }

        public Guid SourcePageWebId { get; set; }

        public string SourcePageWebUrl { get; set; }

        public string SourcePageServerRelativeUrl { get; set; }

        public Guid SourceListWebId { get; set; }

        public Guid SourceListId { get; set; }

        public Guid? SourceViewId { get; set; }

        public string SourceListServerRelativeUrl { get; set; }

        public string SourceTitleUrl { get; set; }

        public string XmlDefinition { get; set; }

        public string JsLink { get; set; }

        public string XslLink { get; set; }

        public string SourceExportSha256 { get; set; }

        public string SourceExportXml { get; set; }
    }

    public sealed class ClassicListWebPartBindingParseResult
    {
        public ClassicListWebPartBindingSnapshot Binding { get; set; }

        public IList<MigrationIssue> Issues { get; set; } = new List<MigrationIssue>();

        public bool IsExecutable => Binding != null && Issues.All(value => value.Severity != MigrationIssueSeverity.Blocker && value.Severity != MigrationIssueSeverity.Error);
    }

    public sealed class ClassicListWebPartTargetMap
    {
        public Guid SourceWebId { get; set; }

        public Guid SourceListId { get; set; }

        public Guid? SourceViewId { get; set; }

        public Guid TargetWebId { get; set; }

        public Guid TargetListId { get; set; }

        public Guid? TargetViewId { get; set; }

        public string TargetListServerRelativeUrl { get; set; }

        public string TargetListAbsoluteUrl { get; set; }

        public string TargetPageServerRelativeUrl { get; set; }

        public IDictionary<string, string> RenderingResourceRewrites { get; set; } = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
    }

    public sealed class RewrittenClassicWebPart
    {
        public Guid SourceWebPartId { get; set; }

        public string ExportXml { get; set; }

        public string ExportSha256 { get; set; }

        public IDictionary<string, string> Replacements { get; set; } = new Dictionary<string, string>();
    }
}
