using PnP.Framework.Migration.Packaging;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;

namespace PnP.Framework.Migration.Schema.Fields
{
    public static class FieldSchemaCanonicalizer
    {
        private static readonly HashSet<string> VolatileAttributes = new HashSet<string>(
            new[] { "List", "WebId", "SourceID", "Version", "ColName", "RowOrdinal" },
            StringComparer.OrdinalIgnoreCase);

        private static readonly HashSet<string> TaxonomyGuidPropertyNames = new HashSet<string>(
            new[] { "SspId", "GroupId", "TermSetId", "AnchorId", "TextField" },
            StringComparer.OrdinalIgnoreCase);

        public static string PortableDigest(string schemaXml)
        {
            return MigrationDigest.ComputeSha256(PortableForm(schemaXml));
        }

        public static string PortableForm(string schemaXml)
        {
            if (string.IsNullOrWhiteSpace(schemaXml))
            {
                throw new ArgumentException("Field schema XML is required.", nameof(schemaXml));
            }

            var document = XDocument.Parse(schemaXml, LoadOptions.None);
            if (document.Root == null)
            {
                throw new ArgumentException("Field schema XML has no root element.", nameof(schemaXml));
            }

            var builder = new StringBuilder();
            AppendElement(document.Root, builder);
            return builder.ToString();
        }

        public static string RewriteForTarget(
            string schemaXml,
            Guid? targetTermStoreId = null,
            Guid? targetTermSetId = null,
            Guid? targetHiddenTextFieldId = null)
        {
            var root = ParseRoot(schemaXml);
            RemoveVolatileAttributes(root);
            SetTaxonomyProperty(root, "SspId", targetTermStoreId);
            SetTaxonomyProperty(root, "TermSetId", targetTermSetId);
            SetTaxonomyProperty(root, "TextField", targetHiddenTextFieldId);
            return root.Document.ToString(SaveOptions.DisableFormatting);
        }

        public static string RewriteLookupForTarget(string schemaXml, Guid targetWebId, Guid targetListId)
        {
            if (targetWebId == Guid.Empty || targetListId == Guid.Empty)
            {
                throw new ArgumentException("Target Web and List IDs are required for a lookup field.");
            }

            var root = ParseRoot(schemaXml);
            RemoveVolatileAttributes(root);
            root.SetAttributeValue("WebId", targetWebId.ToString("D"));
            root.SetAttributeValue("List", $"{{{targetListId:D}}}");
            return root.Document.ToString(SaveOptions.DisableFormatting);
        }

        private static XElement ParseRoot(string schemaXml)
        {
            if (string.IsNullOrWhiteSpace(schemaXml))
            {
                throw new ArgumentException("Field schema XML is required.", nameof(schemaXml));
            }

            var document = XDocument.Parse(schemaXml, LoadOptions.None);
            return document.Root ?? throw new ArgumentException("Field schema XML has no root element.", nameof(schemaXml));
        }

        private static void RemoveVolatileAttributes(XElement root)
        {
            foreach (var attribute in root.Attributes().Where(item => VolatileAttributes.Contains(item.Name.LocalName)).ToArray())
            {
                attribute.Remove();
            }
        }

        private static void SetTaxonomyProperty(XElement root, string name, Guid? value)
        {
            if (!value.HasValue)
            {
                return;
            }

            var property = root.Descendants().FirstOrDefault(element =>
                string.Equals(element.Name.LocalName, "Property", StringComparison.OrdinalIgnoreCase)
                && string.Equals(element.Elements().FirstOrDefault(child => string.Equals(child.Name.LocalName, "Name", StringComparison.OrdinalIgnoreCase))?.Value, name, StringComparison.OrdinalIgnoreCase));
            var propertyValue = property?.Elements().FirstOrDefault(child => string.Equals(child.Name.LocalName, "Value", StringComparison.OrdinalIgnoreCase));
            if (propertyValue != null)
            {
                propertyValue.Value = value.Value.ToString("D");
            }
        }

        private static void AppendElement(XElement element, StringBuilder builder)
        {
            builder.Append('<').Append(ExpandedName(element.Name));
            foreach (var attribute in element.Attributes()
                         .Where(item => !item.IsNamespaceDeclaration)
                         .Where(item => !VolatileAttributes.Contains(item.Name.LocalName))
                         .OrderBy(item => ExpandedName(item.Name), StringComparer.Ordinal))
            {
                var value = attribute.Value;
                if (string.Equals(attribute.Name.NamespaceName, "http://www.w3.org/2001/XMLSchema-instance", StringComparison.Ordinal)
                    && string.Equals(attribute.Name.LocalName, "type", StringComparison.OrdinalIgnoreCase)
                    && value.IndexOf(':') >= 0)
                {
                    value = value.Substring(value.IndexOf(':') + 1);
                }

                builder.Append('|').Append(ExpandedName(attribute.Name)).Append('=').Append(value.Trim());
            }

            builder.Append('>');
            foreach (var node in element.Nodes())
            {
                var child = node as XElement;
                if (child != null && !IsServerGeneratedTaxonomyDefault(child))
                {
                    AppendElement(child, builder);
                    continue;
                }

                var text = node as XText;
                if (text != null && !string.IsNullOrWhiteSpace(text.Value))
                {
                    builder.Append(NormalizeTextValue(element, text.Value.Trim()));
                }
            }

            builder.Append("</").Append(ExpandedName(element.Name)).Append('>');
        }

        private static string ExpandedName(XName name)
        {
            return string.IsNullOrEmpty(name.NamespaceName) ? name.LocalName : $"{{{name.NamespaceName}}}{name.LocalName}";
        }

        private static bool IsServerGeneratedTaxonomyDefault(XElement element)
        {
            if (!string.Equals(element.Name.LocalName, "Property", StringComparison.OrdinalIgnoreCase))
            {
                return false;
            }

            var name = element.Elements().FirstOrDefault(child => string.Equals(child.Name.LocalName, "Name", StringComparison.OrdinalIgnoreCase))?.Value;
            if (!string.Equals(name, "IsDocTagsEnabled", StringComparison.Ordinal)
                && !string.Equals(name, "IsEnhancedImageTaggingEnabled", StringComparison.Ordinal))
            {
                return false;
            }

            var value = element.Elements().FirstOrDefault(child => string.Equals(child.Name.LocalName, "Value", StringComparison.OrdinalIgnoreCase))?.Value;
            return string.Equals(value?.Trim(), "false", StringComparison.OrdinalIgnoreCase);
        }

        private static string NormalizeTextValue(XElement element, string value)
        {
            if (!string.Equals(element.Name.LocalName, "Value", StringComparison.OrdinalIgnoreCase)
                || element.Parent == null
                || !string.Equals(element.Parent.Name.LocalName, "Property", StringComparison.OrdinalIgnoreCase))
            {
                return value;
            }

            var name = element.Parent.Elements().FirstOrDefault(child => string.Equals(child.Name.LocalName, "Name", StringComparison.OrdinalIgnoreCase))?.Value;
            Guid id;
            return TaxonomyGuidPropertyNames.Contains(name ?? string.Empty)
                && Guid.TryParse(value.Trim().Trim('{', '}'), out id)
                    ? id.ToString("D")
                    : value;
        }
    }
}
