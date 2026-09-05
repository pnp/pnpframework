using PnP.Framework.Migration.Pages.Publishing.Capture;
using PnP.Framework.Migration.Pages.Publishing.Planning;
using System.Linq;
using static PnP.Framework.Migration.Pages.Publishing.Reporting.Sections.MigrationReportSectionFormatter;

namespace PnP.Framework.Migration.Pages.Publishing.Reporting.Sections
{
    internal static class ClassicWebPartMigrationReportSection
    {
        public static void Append(MarkdownReportWriter writer, PublishingPageCaptureBundle snapshot, PublishingPageMigrationPlan plan)
        {
            writer.Table($"Shared Web Parts ({snapshot.WebParts.Count})",
                new[] { "ID", "Title", "Zone", "Index", "Hidden", "Export SHA-256", "Export XML" },
                snapshot.WebParts.Select(item => Row(
                    item.Id,
                    item.Title,
                    item.ZoneId,
                    item.ZoneIndex,
                    item.Hidden,
                    item.ExportSha256,
                    Summarize(item.ExportXml))));
            writer.Table($"List-bound Web Part bindings ({snapshot.ListWebPartBindings.Count})",
                new[] { "Web Part", "Type / title", "Page Web / page", "List Web / List / View", "List path / TitleUrl", "Rendering", "XmlDefinition", "Source export" },
                snapshot.ListWebPartBindings.OrderBy(value => value.SourceWebPartId).Select(value => Row(
                    value.SourceWebPartId,
                    $"type={Format(value.TypeName)}; title={Format(value.Title)}",
                    $"web={value.SourcePageWebId:D}; url={Format(value.SourcePageWebUrl)}; page={Format(value.SourcePageServerRelativeUrl)}",
                    $"web={value.SourceListWebId:D}; list={value.SourceListId:D}; view={Format(value.SourceViewId)}",
                    $"path={Format(value.SourceListServerRelativeUrl)}; titleUrl={Format(value.SourceTitleUrl)}",
                    $"jsLink={Format(value.JsLink)}; xslLink={Format(value.XslLink)}",
                    Summarize(value.XmlDefinition),
                    $"sha256={Format(value.SourceExportSha256)}; xml={Summarize(value.SourceExportXml)}")));
            writer.Table($"Web Part plan actions ({plan.WebPartActions.Count})",
                new[] { "Web Part", "Disposition", "Source List Web / List / View", "Target Web / List path", "Reason" },
                plan.WebPartActions.OrderBy(value => value.SourceWebPartId).Select(value => Row(
                    value.SourceWebPartId,
                    value.Disposition,
                    $"web={Format(value.SourceListWebId)}; list={Format(value.SourceListId)}; view={Format(value.SourceViewId)}",
                    $"web={Format(value.TargetWebUrl)}; listPath={Format(value.TargetListServerRelativeUrl)}",
                    value.Reason)));
        }
    }
}
