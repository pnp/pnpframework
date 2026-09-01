using PnP.Framework.Migration.Pages.Packaging;
using PnP.Framework.Migration.Pages.Publishing.Capture;
using PnP.Framework.Migration.Pages.Publishing.Planning;
using System;

namespace PnP.Framework.Migration.Pages.Publishing.Packaging
{
    public static class PublishingPageDigest
    {
        public static string ComputeSnapshotDigest(PublishingPageCaptureBundle snapshot)
        {
            if (snapshot == null)
            {
                throw new ArgumentNullException(nameof(snapshot));
            }

            return PageDigest.ComputeSha256(PublishingPagePackageSerializer.SerializeCanonical(snapshot));
        }

        public static string ComputePlanDigest(PublishingPageMigrationPlan plan)
        {
            if (plan == null)
            {
                throw new ArgumentNullException(nameof(plan));
            }

            return PageDigest.ComputeSha256(PublishingPagePackageSerializer.SerializeCanonical(plan));
        }

        public static string ComputeSha256(string value)
        {
            return PageDigest.ComputeSha256(value);
        }

        public static string ComputeSha256(byte[] value)
        {
            return PageDigest.ComputeSha256(value);
        }
    }
}
