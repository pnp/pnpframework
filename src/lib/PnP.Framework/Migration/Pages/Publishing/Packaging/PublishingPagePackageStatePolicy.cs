using PnP.Framework.Migration.Pages.Ingredients;
using PnP.Framework.Migration.Pages.Publishing.Planning;
using System;

namespace PnP.Framework.Migration.Pages.Publishing.Packaging
{
    internal static class PublishingPagePackageStatePolicy
    {
        public static PublishingPagePackageState Derive(PublishingPageMigrationPlan plan)
        {
            if (plan == null)
            {
                throw new ArgumentNullException(nameof(plan));
            }
            if (plan.IsExecutable)
            {
                return PublishingPagePackageState.ApprovalReady;
            }
            if (plan.MigrationOutcome == PageMigrationOutcome.AuthorizationBlocked)
            {
                return PublishingPagePackageState.AuthorizationBlocked;
            }
            if (plan.MigrationOutcome == PageMigrationOutcome.MitigationPending
                || ((plan.Blockers?.Count ?? 0) > 0
                    && plan.MigrationOutcome != PageMigrationOutcome.Invalid
                    && plan.MigrationOutcome != PageMigrationOutcome.Unknown
                    && plan.MigrationOutcome != PageMigrationOutcome.Blocked))
            {
                return PublishingPagePackageState.MitigationPending;
            }
            return PublishingPagePackageState.Invalid;
        }
    }
}
