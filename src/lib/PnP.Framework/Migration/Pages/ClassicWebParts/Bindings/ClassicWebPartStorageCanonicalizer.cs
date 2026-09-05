using System;
using System.Linq;
using System.Xml.Linq;

namespace PnP.Framework.Migration.Pages.ClassicWebParts.Bindings
{
    internal static class ClassicWebPartStorageCanonicalizer
    {
        public static string CanonicalizeListBoundExport(string exportXml)
        {
            var document = XDocument.Parse(exportXml ?? string.Empty);
            foreach (var property in document.Descendants().Where(IsProperty))
            {
                var name = (string)property.Attribute("name") ?? string.Empty;
                if (IsGuidProperty(name))
                {
                    property.Value = NormalizeGuid(
                        property.Value,
                        string.Equals(name, "ListName", StringComparison.OrdinalIgnoreCase));
                }
                else if (string.Equals(name, "XmlDefinition", StringComparison.OrdinalIgnoreCase))
                {
                    property.Value = CanonicalizeViewDefinition(property.Value);
                }
                else if (string.Equals(name, "NoDefaultStyle", StringComparison.OrdinalIgnoreCase)
                    && string.IsNullOrEmpty(property.Value))
                {
                    property.SetAttributeValue("null", null);
                }
            }

            return document.ToString(SaveOptions.DisableFormatting);
        }

        private static bool IsProperty(XElement value) =>
            string.Equals(value.Name.LocalName, "property", StringComparison.OrdinalIgnoreCase);

        private static bool IsGuidProperty(string name) =>
            string.Equals(name, "WebId", StringComparison.OrdinalIgnoreCase)
            || string.Equals(name, "ListId", StringComparison.OrdinalIgnoreCase)
            || string.Equals(name, "ListName", StringComparison.OrdinalIgnoreCase)
            || string.Equals(name, "ViewGuid", StringComparison.OrdinalIgnoreCase);

        private static string NormalizeGuid(string value, bool braces)
        {
            if (!Guid.TryParse((value ?? string.Empty).Trim().Trim('{', '}'), out var parsed))
            {
                return value ?? string.Empty;
            }

            var normalized = parsed.ToString("D");
            return braces ? "{" + normalized + "}" : normalized;
        }

        private static string CanonicalizeViewDefinition(string value)
        {
            if (string.IsNullOrWhiteSpace(value))
            {
                return value ?? string.Empty;
            }

            var view = XDocument.Parse(value);
            var name = view.Root?.Attribute("Name");
            if (name != null && Guid.TryParse(name.Value.Trim().Trim('{', '}'), out _))
            {
                name.Value = "{runtime-view-identity}";
            }
            var imageUrl = view.Root?.Attribute("ImageUrl");
            var layoutsIndex = imageUrl?.Value.LastIndexOf("/_layouts/", StringComparison.OrdinalIgnoreCase) ?? -1;
            if (imageUrl != null && layoutsIndex >= 0)
            {
                // SharePoint stores built-in View icons root-relative. Accept
                // absolute source/target forms and the legacy composer defect
                // that concatenated an origin twice before the same _layouts
                // path; everything before the final built-in path is storage
                // noise rather than authored content.
                imageUrl.Value = imageUrl.Value.Substring(layoutsIndex);
            }

            return view.ToString(SaveOptions.DisableFormatting);
        }
    }
}
