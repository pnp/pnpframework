using PnP.Framework.Migration.Diagnostics;
using System.Collections.Generic;
using System.Linq;

namespace PnP.Framework.Migration.Pages.Ingredients
{
    public sealed class PageIngredientPlanEvaluation
    {
        public PageMigrationOutcome Outcome { get; set; }

        public IList<MigrationIssue> Issues { get; set; } = new List<MigrationIssue>();

        public bool IsExecutable => Outcome != PageMigrationOutcome.Blocked
            && Outcome != PageMigrationOutcome.Unknown
            && Issues.All(value => value.Severity != MigrationIssueSeverity.Blocker && value.Severity != MigrationIssueSeverity.Error);
    }
}
