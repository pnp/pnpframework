using System;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text;

namespace PnP.Framework.Migration.Packaging
{
    public static class MigrationDigest
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
                return string.Concat(algorithm.ComputeHash(value).Select(item => item.ToString("x2", CultureInfo.InvariantCulture)));
            }
        }

        public static string ComputeSha256(Stream value)
        {
            if (value == null)
            {
                throw new ArgumentNullException(nameof(value));
            }

            using (var algorithm = SHA256.Create())
            {
                return string.Concat(algorithm.ComputeHash(value).Select(item => item.ToString("x2", CultureInfo.InvariantCulture)));
            }
        }
    }
}
