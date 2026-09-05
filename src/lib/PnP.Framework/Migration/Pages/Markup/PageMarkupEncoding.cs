using System;
using System.Text;

namespace PnP.Framework.Migration.Pages.Markup
{
    internal static class PageMarkupEncoding
    {
        public static string Decode(byte[] bytes)
        {
            if (bytes == null)
            {
                throw new ArgumentNullException(nameof(bytes));
            }

            var encoding = GetEncoding(bytes, out var preambleLength);
            return encoding.GetString(bytes, preambleLength, bytes.Length - preambleLength);
        }

        private static Encoding GetEncoding(byte[] bytes, out int preambleLength)
        {
            if (bytes.Length >= 2 && bytes[0] == 0xff && bytes[1] == 0xfe)
            {
                preambleLength = 2;
                return Encoding.Unicode;
            }

            if (bytes.Length >= 2 && bytes[0] == 0xfe && bytes[1] == 0xff)
            {
                preambleLength = 2;
                return Encoding.BigEndianUnicode;
            }

            if (bytes.Length >= 3 && bytes[0] == 0xef && bytes[1] == 0xbb && bytes[2] == 0xbf)
            {
                preambleLength = 3;
                return new UTF8Encoding(true);
            }

            preambleLength = 0;
            return new UTF8Encoding(false);
        }
    }
}
