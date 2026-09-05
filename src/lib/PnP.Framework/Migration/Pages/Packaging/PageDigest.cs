using System;
using PnP.Framework.Migration.Packaging;

namespace PnP.Framework.Migration.Pages.Packaging
{
    internal static class PageDigest
    {
        public static string ComputeSha256(string value)
        {
            if (value == null)
            {
                throw new ArgumentNullException(nameof(value));
            }

            return MigrationDigest.ComputeSha256(value);
        }

        public static string ComputeSha256(byte[] value)
        {
            if (value == null)
            {
                throw new ArgumentNullException(nameof(value));
            }

            return MigrationDigest.ComputeSha256(value);
        }
    }
}
