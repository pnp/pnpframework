using PnP.Framework.Migration.Pages.Publishing.Capture;
using PnP.Framework.Migration.Pages.Publishing.Planning;
using System.Collections.Generic;
using System.Linq;

namespace PnP.Framework.Migration.Pages.Publishing.Reporting
{
    internal static class PublishingPagePlanReportFactory
    {
        public static PublishingPageMigrationReport Create(
            PublishingPageCaptureBundle snapshot,
            PublishingPageMigrationPlan plan)
        {
            return new PublishingPageMigrationReport
            {
                Summary = plan.IsExecutable
                    ? "Source export and target analysis completed. Import requires explicit approval of the sealed plan digest."
                    : "The package is sealed for review but cannot be imported until every blocker is resolved and a new plan is generated.",
                CapturedIngredients = new List<string>
                {
                    "Page/file/list item identity and source stability fence",
                    "Exact source ASPX artifact and Page directive",
                    "CLR runtime evidence and non-exclusive product profile signals",
                    $"All {snapshot.Fields.Count} source Pages-library field definitions and returned values",
                    $"{snapshot.WebParts.Count} shared Web Part export(s) with zone placement",
                    $"{snapshot.Dependencies.Count} authored dependency/link snapshot(s)",
                    "Canonical ingredient dependency graph, per-ingredient capability and disposition",
                    "Page security inheritance and source lifecycle evidence",
                    "Target publishing library, versioning, lifecycle, field, layout, and create-only probes"
                },
                Blockers = plan.Blockers.ToList(),
                Warnings = plan.Warnings.ToList()
            };
        }
    }
}
