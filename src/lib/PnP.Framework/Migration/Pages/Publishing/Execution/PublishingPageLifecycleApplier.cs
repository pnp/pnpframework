using Microsoft.SharePoint.Client;
using PnP.Framework.Migration.Execution;
using PnP.Framework.Migration.Pages.Publishing.Lifecycle;
using PnP.Framework.Migration.Pages.Publishing.Packaging;
using System.Collections.Generic;

namespace PnP.Framework.Migration.Pages.Publishing.Execution
{
    internal static class PublishingPageLifecycleApplier
    {
        public static void Apply(
            ClientContext context,
            PublishingPageMigrationPackage package,
            PublishingPageWriteResult result,
            bool applyApprovedLifecycle,
            MigrationExecutionRecorder recorder,
            ICollection<string> warnings)
        {
            context.Load(result.TargetFile, file => file.CheckOutType);
            context.ExecuteQueryRetry();
            if (result.TargetFile.CheckOutType != CheckOutType.None)
            {
                var checkinType = applyApprovedLifecycle
                    && package.Plan.TargetLifecycle == PublishingPageTargetLifecycle.Published
                    && !result.PlannedFieldFailure
                    ? CheckinType.MajorCheckIn
                    : CheckinType.MinorCheckIn;
                recorder.Execute("page.checkin", $"Check in the page as {checkinType}.", () =>
                {
                    result.TargetFile.CheckIn("PnP publishing-page import", checkinType);
                    context.ExecuteQueryRetry();
                });
            }
            else
            {
                recorder.RecordAlreadySatisfied("page.checkin", "The target page is already checked in.");
            }

            if (!applyApprovedLifecycle
                || package.Plan.TargetLifecycle != PublishingPageTargetLifecycle.Published
                || result.PlannedFieldFailure)
            {
                recorder.RecordAlreadySatisfied(
                    "page.publish",
                    !applyApprovedLifecycle
                        ? "The lifecycle ingredient is outside the admitted execution frontier; the page is left Draft."
                        : "The approved lifecycle is Draft, or a planned field update failed; no publish action was performed.");
                if (result.PlannedFieldFailure)
                {
                    warnings.Add("One or more planned field updates failed. The page was not published.");
                }

                return;
            }

            recorder.Execute("page.publish", "Publish the page because the approved lifecycle is Published.", () =>
            {
                result.TargetFile.Publish("PnP publishing-page import");
                context.ExecuteQueryRetry();
            });
            if (result.PagesLibrary.EnableModeration)
            {
                recorder.Execute("page.approve", "Approve the published page because moderation is enabled.", () =>
                {
                    result.TargetFile.Approve("PnP publishing-page import");
                    context.ExecuteQueryRetry();
                });
            }
            else
            {
                recorder.RecordAlreadySatisfied("page.approve", "The target Pages library does not require moderation approval.");
            }
        }
    }
}
