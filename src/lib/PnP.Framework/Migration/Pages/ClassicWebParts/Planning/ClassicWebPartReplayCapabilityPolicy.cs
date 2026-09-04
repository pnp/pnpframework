using System;
using System.Linq;
using System.Xml.Linq;

namespace PnP.Framework.Migration.Pages.ClassicWebParts.Planning
{
    internal static class ClassicWebPartReplayCapabilityPolicy
    {
        public static string GetBlocker(string exportXml)
        {
            if (string.IsNullOrWhiteSpace(exportXml))
            {
                return "the shared Web Part export is empty";
            }

            XDocument document;
            try
            {
                document = XDocument.Parse(exportXml, LoadOptions.None);
            }
            catch (Exception exception) when (exception is System.Xml.XmlException || exception is ArgumentException)
            {
                return $"the shared Web Part export is not valid XML ({exception.Message})";
            }

            var typeName = document
                .Descendants()
                .Where(element => string.Equals(element.Name.LocalName, "type", StringComparison.OrdinalIgnoreCase))
                .Select(element => (string)element.Attribute("name"))
                .FirstOrDefault(value => !string.IsNullOrWhiteSpace(value));
            if (string.IsNullOrWhiteSpace(typeName))
            {
                typeName = document
                    .Descendants()
                    .Where(element => string.Equals(element.Name.LocalName, "TypeName", StringComparison.OrdinalIgnoreCase))
                    .Select(element => element.Value?.Trim())
                    .FirstOrDefault(value => !string.IsNullOrWhiteSpace(value));
            }
            if (string.IsNullOrWhiteSpace(typeName))
            {
                return "the shared Web Part export does not declare a v2 or v3 type";
            }

            if (typeName.EndsWith(".ErrorWebPart", StringComparison.OrdinalIgnoreCase))
            {
                return $"type '{typeName}' represents a Web Part that is already unavailable on the source page";
            }

            if (typeName.EndsWith(".RSSAggregatorWebPart", StringComparison.OrdinalIgnoreCase))
            {
                return $"type '{typeName}' is not supported by the current deterministic Publishing Page importer and must be replaced or explicitly mapped";
            }

            var properties = document
                .Descendants()
                .Where(element => string.Equals(element.Name.LocalName, "property", StringComparison.OrdinalIgnoreCase))
                .ToArray();
            var hasSourceListBinding = properties.Any(property =>
            {
                var name = (string)property.Attribute("name");
                if (!string.Equals(name, "ListId", StringComparison.OrdinalIgnoreCase)
                    && !string.Equals(name, "ListName", StringComparison.OrdinalIgnoreCase))
                {
                    return false;
                }

                var value = property.Value?.Trim().Trim('{', '}');
                return !string.IsNullOrWhiteSpace(value)
                    && !string.Equals(value, Guid.Empty.ToString(), StringComparison.OrdinalIgnoreCase);
            });
            if (hasSourceListBinding)
            {
                return null;
            }

            if (typeName.EndsWith(".XsltListViewWebPart", StringComparison.OrdinalIgnoreCase)
                || typeName.EndsWith(".ListViewWebPart", StringComparison.OrdinalIgnoreCase))
            {
                return $"type '{typeName}' is list-oriented but its export has no parseable source ListId/ListName binding";
            }

            return null;
        }
    }
}
