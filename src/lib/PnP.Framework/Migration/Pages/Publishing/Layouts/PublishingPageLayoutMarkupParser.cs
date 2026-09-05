using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Text.RegularExpressions;
using PnP.Framework.Migration.Pages.Markup;

namespace PnP.Framework.Migration.Pages.Publishing.Layouts
{
    internal static class PublishingPageLayoutMarkupParser
    {
        private static readonly Regex RegisterDirectivePattern = new Regex(
            "<%@\\s*Register\\s+(?<attributes>.*?)%>",
            RegexOptions.IgnoreCase | RegexOptions.Singleline | RegexOptions.CultureInvariant | RegexOptions.Compiled);

        private static readonly Regex ControlPattern = new Regex(
            "<(?<prefix>[A-Za-z_][A-Za-z0-9_]*)\\:(?<name>[A-Za-z_][A-Za-z0-9_]*)(?<attributes>(?:[^>\\\"']|\\\"[^\\\"]*\\\"|'[^']*')*)/?>",
            RegexOptions.IgnoreCase | RegexOptions.Singleline | RegexOptions.CultureInvariant | RegexOptions.Compiled);

        private static readonly Regex AttributePattern = new Regex(
            "(?<name>[A-Za-z_][A-Za-z0-9_:-]*)\\s*=\\s*(?:\"(?<value>[^\"]*)\"|'(?<value>[^']*)')",
            RegexOptions.IgnoreCase | RegexOptions.CultureInvariant | RegexOptions.Compiled);

        private static readonly Regex ResourcePattern = new Regex(
            "(?<attribute>src|href|poster|data)\\s*=\\s*(?:\"(?<value>[^\"]+)\"|'(?<value>[^']+)')",
            RegexOptions.IgnoreCase | RegexOptions.CultureInvariant | RegexOptions.Compiled);

        private static readonly Regex CssResourcePattern = new Regex(
            "url\\(\\s*(?:\"(?<value>[^\"]+)\"|'(?<value>[^']+)'|(?<value>[^)\\s]+))\\s*\\)",
            RegexOptions.IgnoreCase | RegexOptions.CultureInvariant | RegexOptions.Compiled);

        private static readonly Regex SpUrlExpressionPattern = new Regex(
            "^<%\\s*\\$SPUrl:(?<value>.*?)%>$",
            RegexOptions.IgnoreCase | RegexOptions.Singleline | RegexOptions.CultureInvariant | RegexOptions.Compiled);

        private static readonly HashSet<string> ControlResourceAttributeNames = new HashSet<string>(
            new[]
            {
                "Src", "Href", "Poster", "Data", "ImageUrl", "NodeImageUrl", "RTLNodeImageUrl",
                "ExpandImageUrl", "ExpandImageUrlRtl", "CollapseImageUrl", "CollapseImageUrlRtl",
                "NoExpandImageUrl", "NavigateUrl", "CssFileLocation", "ScriptFile"
            },
            StringComparer.OrdinalIgnoreCase);

        public static PublishingPageLayoutMarkup Parse(string markup)
        {
            if (markup == null)
            {
                throw new ArgumentNullException(nameof(markup));
            }

            var registrations = ParseRegistrations(markup);
            var controls = ParseControls(markup);
            return new PublishingPageLayoutMarkup
            {
                PageDirective = PageDirectiveParser.Parse(markup),
                Registrations = registrations,
                Controls = controls,
                Zones = controls
                    .Where(item => string.Equals(item.ControlName, "WebPartZone", StringComparison.OrdinalIgnoreCase))
                    .Where(item => !string.IsNullOrWhiteSpace(item.Id))
                    .GroupBy(item => item.Id, StringComparer.OrdinalIgnoreCase)
                    .Select(group => new PublishingPageLayoutZone { Id = group.Key })
                    .OrderBy(item => item.Id, StringComparer.OrdinalIgnoreCase)
                    .ToList(),
                RequiredFieldIdentifiers = controls
                    .Select(item => item.FieldName)
                    .Where(item => !string.IsNullOrWhiteSpace(item))
                    .Concat(new[] { "Title", "PublishingPageContent" })
                    .Distinct(StringComparer.OrdinalIgnoreCase)
                    .OrderBy(item => item, StringComparer.OrdinalIgnoreCase)
                    .ToList(),
                ResourceReferences = ParseResourceReferences(markup)
            };
        }

        private static IList<PublishingPageLayoutRegistration> ParseRegistrations(string markup)
        {
            return RegisterDirectivePattern.Matches(markup).Cast<Match>()
                .Select(match => Attributes(match.Groups["attributes"].Value))
                .Select(attributes => new PublishingPageLayoutRegistration
                {
                    TagPrefix = Value(attributes, "TagPrefix"),
                    Namespace = Value(attributes, "Namespace"),
                    Assembly = Value(attributes, "Assembly")
                })
                .Where(item => !string.IsNullOrWhiteSpace(item.TagPrefix)
                    || !string.IsNullOrWhiteSpace(item.Namespace)
                    || !string.IsNullOrWhiteSpace(item.Assembly))
                .GroupBy(item => $"{item.TagPrefix}|{item.Namespace}|{item.Assembly}", StringComparer.Ordinal)
                .Select(group => group.First())
                .OrderBy(item => item.TagPrefix, StringComparer.OrdinalIgnoreCase)
                .ThenBy(item => item.Namespace, StringComparer.Ordinal)
                .ToList();
        }

        private static IList<PublishingPageLayoutControl> ParseControls(string markup)
        {
            return ControlPattern.Matches(markup).Cast<Match>()
                .Select(match =>
                {
                    var attributes = Attributes(match.Groups["attributes"].Value);
                    return new PublishingPageLayoutControl
                    {
                        TagPrefix = match.Groups["prefix"].Value,
                        ControlName = match.Groups["name"].Value,
                        Id = Value(attributes, "ID"),
                        FieldName = Value(attributes, "FieldName")
                    };
                })
                .GroupBy(item => $"{item.TagPrefix}|{item.ControlName}|{item.Id}|{item.FieldName}", StringComparer.OrdinalIgnoreCase)
                .Select(group => group.First())
                .OrderBy(item => item.TagPrefix, StringComparer.OrdinalIgnoreCase)
                .ThenBy(item => item.ControlName, StringComparer.OrdinalIgnoreCase)
                .ThenBy(item => item.Id, StringComparer.OrdinalIgnoreCase)
                .ToList();
        }

        private static IList<PublishingPageLayoutResourceReference> ParseResourceReferences(string markup)
        {
            var representations = new[] { markup, WebUtility.HtmlDecode(markup) }.Distinct(StringComparer.Ordinal).ToArray();
            var html = representations.SelectMany(value => ResourcePattern.Matches(value).Cast<Match>().Select(match =>
                Reference(match.Groups["attribute"].Value.ToLowerInvariant(), match.Groups["value"].Value)));
            var css = representations.SelectMany(value => CssResourcePattern.Matches(value).Cast<Match>().Select(match =>
                Reference("css-url", match.Groups["value"].Value)));
            var controls = ControlPattern.Matches(markup).Cast<Match>().SelectMany(match =>
                ControlResourceReferences(match.Groups["prefix"].Value, match.Groups["name"].Value, Attributes(match.Groups["attributes"].Value)));
            return html.Concat(css).Concat(controls)
                .Where(item => !string.IsNullOrWhiteSpace(item.Value))
                .Where(item => !item.Value.StartsWith("javascript:", StringComparison.OrdinalIgnoreCase))
                .Where(item => !item.Value.StartsWith("data:", StringComparison.OrdinalIgnoreCase))
                .Where(item => !item.Value.StartsWith("#", StringComparison.Ordinal))
                .GroupBy(item => $"{item.Attribute}|{item.Value}", StringComparer.OrdinalIgnoreCase)
                .Select(group => group.First())
                .OrderBy(item => item.Attribute, StringComparer.Ordinal)
                .ThenBy(item => item.Value, StringComparer.OrdinalIgnoreCase)
                .ToList();
        }

        private static IEnumerable<PublishingPageLayoutResourceReference> ControlResourceReferences(
            string prefix,
            string controlName,
            IDictionary<string, string> attributes)
        {
            foreach (var attribute in attributes)
            {
                var isRegistrationResource =
                    (string.Equals(controlName, "CssRegistration", StringComparison.OrdinalIgnoreCase)
                        && (string.Equals(attribute.Key, "Name", StringComparison.OrdinalIgnoreCase)
                            || string.Equals(attribute.Key, "After", StringComparison.OrdinalIgnoreCase)))
                    || (string.Equals(controlName, "ScriptLink", StringComparison.OrdinalIgnoreCase)
                        && string.Equals(attribute.Key, "Name", StringComparison.OrdinalIgnoreCase));
                if (!isRegistrationResource && !ControlResourceAttributeNames.Contains(attribute.Key))
                {
                    continue;
                }

                var value = NormalizeResourceReference(attribute.Value);
                if (LooksLikeResourceReference(value))
                {
                    yield return Reference($"control:{prefix}:{controlName}:{attribute.Key}", value);
                }
            }
        }

        private static PublishingPageLayoutResourceReference Reference(string attribute, string value)
        {
            return new PublishingPageLayoutResourceReference
            {
                Attribute = attribute,
                Value = NormalizeResourceReference(value)
            };
        }

        private static string NormalizeResourceReference(string value)
        {
            var decoded = WebUtility.HtmlDecode(value ?? string.Empty).Trim();
            var match = SpUrlExpressionPattern.Match(decoded);
            return match.Success ? match.Groups["value"].Value.Trim() : decoded;
        }

        private static bool LooksLikeResourceReference(string value)
        {
            return !string.IsNullOrWhiteSpace(value)
                && (value.StartsWith("/", StringComparison.Ordinal)
                    || value.StartsWith("~/", StringComparison.Ordinal)
                    || value.StartsWith("~site/", StringComparison.OrdinalIgnoreCase)
                    || value.StartsWith("~sitecollection/", StringComparison.OrdinalIgnoreCase)
                    || value.StartsWith("http://", StringComparison.OrdinalIgnoreCase)
                    || value.StartsWith("https://", StringComparison.OrdinalIgnoreCase));
        }

        private static Dictionary<string, string> Attributes(string text)
        {
            return AttributePattern.Matches(text ?? string.Empty).Cast<Match>()
                .GroupBy(match => match.Groups["name"].Value, StringComparer.OrdinalIgnoreCase)
                .ToDictionary(
                    group => group.Key,
                    group => WebUtility.HtmlDecode(group.Last().Groups["value"].Value),
                    StringComparer.OrdinalIgnoreCase);
        }

        private static string Value(IDictionary<string, string> values, string key)
        {
            string value;
            return values.TryGetValue(key, out value) ? value : string.Empty;
        }
    }
}
