using PnP.Framework.Migration.PublishingPages.Capture;
using PnP.Framework.Migration.PublishingPages.Planning;
using System;
using System.Globalization;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;

namespace PnP.Framework.Migration.PublishingPages.Packaging
{
    public static class PublishingPageDigest
    {
        public static string ComputeSnapshotDigest(PublishingPageCaptureBundle snapshot)
        {
            if (snapshot == null)
            {
                throw new ArgumentNullException(nameof(snapshot));
            }

            return ComputeSha256(PublishingPagePackageSerializer.SerializeCanonical(snapshot));
        }

        public static string ComputePlanDigest(PublishingPageMigrationPlan plan)
        {
            if (plan == null)
            {
                throw new ArgumentNullException(nameof(plan));
            }

            return ComputeSha256(PublishingPagePackageSerializer.SerializeCanonical(plan));
        }

        public static string ComputeSha256(string value)
        {
            if (value == null)
            {
                throw new ArgumentNullException(nameof(value));
            }

            return ComputeSha256(Encoding.UTF8.GetBytes(value));
        }

        public static string ComputeSha256(byte[] value)
        {
            if (value == null)
            {
                throw new ArgumentNullException(nameof(value));
            }

            using (var algorithm = SHA256.Create())
            {
                var digest = algorithm.ComputeHash(value);
                return string.Concat(digest.Select(item => item.ToString("x2", CultureInfo.InvariantCulture)));
            }
        }
    }
}
