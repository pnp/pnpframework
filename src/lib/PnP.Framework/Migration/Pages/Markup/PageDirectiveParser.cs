using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Text.RegularExpressions;

namespace PnP.Framework.Migration.Pages.Markup
{
    internal static class PageDirectiveParser
    {
        private static readonly Regex PageDirectivePattern = new Regex(
            "<%@\\s*Page\\s+(?<attributes>.*?)%>",
            RegexOptions.IgnoreCase | RegexOptions.Singleline | RegexOptions.CultureInvariant | RegexOptions.Compiled);

        private static readonly Regex AttributePattern = new Regex(
            "(?<name>[A-Za-z_][A-Za-z0-9_:-]*)\\s*=\\s*(?:\"(?<value>[^\"]*)\"|'(?<value>[^']*)')",
            RegexOptions.IgnoreCase | RegexOptions.CultureInvariant | RegexOptions.Compiled);

        public static PageDirectiveSnapshot Parse(string markup)
        {
            if (markup == null)
            {
                throw new ArgumentNullException(nameof(markup));
            }

            var match = PageDirectivePattern.Match(markup);
            if (!match.Success)
            {
                return null;
            }

            var attributes = AttributePattern.Matches(match.Groups["attributes"].Value)
                .Cast<Match>()
                .Select(value => new PageDirectiveAttribute
                {
                    Name = value.Groups["name"].Value,
                    Value = WebUtility.HtmlDecode(value.Groups["value"].Value)
                })
                .GroupBy(value => value.Name, StringComparer.OrdinalIgnoreCase)
                .Select(group => group.Last())
                .OrderBy(value => value.Name, StringComparer.OrdinalIgnoreCase)
                .ToList();

            return new PageDirectiveSnapshot
            {
                Inherits = Value(attributes, "Inherits"),
                MasterPageFile = Value(attributes, "MasterPageFile"),
                Language = Value(attributes, "Language"),
                CodeBehind = Value(attributes, "CodeBehind"),
                CodeFile = Value(attributes, "CodeFile"),
                Attributes = attributes
            };
        }

        private static string Value(IEnumerable<PageDirectiveAttribute> attributes, string name)
        {
            return attributes.FirstOrDefault(value => string.Equals(value.Name, name, StringComparison.OrdinalIgnoreCase))?.Value;
        }
    }
}
