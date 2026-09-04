using PnP.Framework.Migration.Diagnostics;
using System.Collections.Generic;
using System.Linq;

namespace PnP.Framework.Migration.Pages.Ingredients
{
    public sealed class PageIngredientPlanEvaluation
    {
        public PageMigrationOutcome Outcome { get; set; }

        public IList<MigrationIssue> Issues { get; set; } = new List<MigrationIssue>();

        public PageIngredientExecutionFrontier ExecutionFrontier { get; set; } = new PageIngredientExecutionFrontier();

        public bool IsExecutable => (Outcome == PageMigrationOutcome.Exact
                || Outcome == PageMigrationOutcome.ExecutableWithTransform
                || Outcome == PageMigrationOutcome.ExecutableWithLoss
                || Outcome == PageMigrationOutcome.PartiallyExecutable)
            && Issues.All(value => value.Severity != MigrationIssueSeverity.Error)
            && (ExecutionFrontier.HasExecutableIngredients
                || Outcome == PageMigrationOutcome.ExecutableWithLoss);
    }
}
