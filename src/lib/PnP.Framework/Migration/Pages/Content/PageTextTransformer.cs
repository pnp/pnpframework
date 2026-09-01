using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

namespace PnP.Framework.Migration.Pages.Content
{
    internal static class PageTextTransformer
    {
        public static string Rewrite(string value, IEnumerable<PageTextReplacement> replacements)
        {
            var result = value ?? string.Empty;
            foreach (var replacement in (replacements ?? Array.Empty<PageTextReplacement>())
                         .Where(item => !string.IsNullOrEmpty(item.Source))
                         .OrderByDescending(item => item.Source.Length))
            {
                result = Regex.Replace(
                    result,
                    Regex.Escape(replacement.Source),
                    _ => replacement.Target ?? string.Empty,
                    RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);
            }

            return result;
        }
    }
}
