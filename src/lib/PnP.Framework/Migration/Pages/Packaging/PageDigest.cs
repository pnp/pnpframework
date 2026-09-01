using System;
using System.Globalization;
using System.Linq;
using System.Security.Cryptography;
using System.Text;

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
