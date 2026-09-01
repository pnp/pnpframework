using System;
using System.Text;

namespace PnP.Framework.Migration.Pages.Publishing.Layouts
{
    internal static class PublishingPageLayoutEncoding
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

        public static byte[] Encode(string value, byte[] originalBytes)
        {
            if (value == null)
            {
                throw new ArgumentNullException(nameof(value));
            }

            var encoding = GetEncoding(originalBytes, out var preambleLength);
            var body = encoding.GetBytes(value);
            if (preambleLength == 0)
            {
                return body;
            }

            var preamble = encoding.GetPreamble();
            var result = new byte[preamble.Length + body.Length];
            Buffer.BlockCopy(preamble, 0, result, 0, preamble.Length);
            Buffer.BlockCopy(body, 0, result, preamble.Length, body.Length);
            return result;
        }

        private static Encoding GetEncoding(byte[] bytes, out int preambleLength)
        {
            if (bytes != null && bytes.Length >= 2 && bytes[0] == 0xff && bytes[1] == 0xfe)
            {
                preambleLength = 2;
                return Encoding.Unicode;
            }

            if (bytes != null && bytes.Length >= 2 && bytes[0] == 0xfe && bytes[1] == 0xff)
            {
                preambleLength = 2;
                return Encoding.BigEndianUnicode;
            }

            if (bytes != null && bytes.Length >= 3 && bytes[0] == 0xef && bytes[1] == 0xbb && bytes[2] == 0xbf)
            {
                preambleLength = 3;
                return new UTF8Encoding(true);
            }

            preambleLength = 0;
            return new UTF8Encoding(false);
        }
    }
}
