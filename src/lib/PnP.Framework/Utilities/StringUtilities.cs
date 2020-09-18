using System.Text.RegularExpressions;

namespace PnP.Framework.Utilities
{
    public static class StringUtilities
    {
        public static string[] Split(this string input, string separator)
        {
            var splitRegex = new Regex(
                Regex.Escape(separator),
                RegexOptions.Singleline | RegexOptions.Compiled
                );

            return splitRegex.Split(input);
        }
    }
}
