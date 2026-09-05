using PnP.Framework.Migration.Packaging;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml.Linq;

namespace PnP.Framework.Migration.Pages.ClassicWebParts.Bindings
{
    public static class ClassicListWebPartRewriter
    {
        public static RewrittenClassicWebPart Rewrite(ClassicListWebPartBindingSnapshot binding, ClassicListWebPartTargetMap target)
        {
            if (binding == null)
            {
                throw new ArgumentNullException(nameof(binding));
            }
            if (target == null)
            {
                throw new ArgumentNullException(nameof(target));
            }
            if (binding.SourceListId != target.SourceListId || binding.SourceListWebId != target.SourceWebId)
            {
                throw new ArgumentException("The target mapping does not match the source Web Part list binding.", nameof(target));
            }

            var document = XDocument.Parse(binding.SourceExportXml, LoadOptions.PreserveWhitespace);
            var replacements = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            SetProperty(document, "WebId", target.TargetWebId.ToString("D"), replacements);
            SetProperty(document, "ListId", target.TargetListId.ToString("D"), replacements);
            SetProperty(document, "ListName", "{" + target.TargetListId.ToString("D") + "}", replacements);
            if (!string.IsNullOrWhiteSpace(binding.SourceTitleUrl))
            {
                SetProperty(document, "TitleUrl", target.TargetListServerRelativeUrl, replacements);
            }
            if (target.TargetViewId.HasValue)
            {
                SetOptionalProperty(document, "ViewGuid", target.TargetViewId.Value.ToString("D"), replacements);
            }

            var definition = FindProperty(document, "XmlDefinition");
            if (definition != null)
            {
                var view = XDocument.Parse(definition.Value, LoadOptions.PreserveWhitespace);
                if (target.TargetViewId.HasValue && view.Root != null)
                {
                    var old = (string)view.Root.Attribute("Name");
                    view.Root.SetAttributeValue("Name", "{" + target.TargetViewId.Value.ToString("D") + "}");
                    replacements["XmlDefinition/View@Name:" + old] = target.TargetViewId.Value.ToString("D");
                }
                if (!string.IsNullOrWhiteSpace(target.TargetPageServerRelativeUrl) && view.Root != null)
                {
                    var old = (string)view.Root.Attribute("Url");
                    view.Root.SetAttributeValue("Url", target.TargetPageServerRelativeUrl);
                    replacements["XmlDefinition/View@Url:" + old] = target.TargetPageServerRelativeUrl;
                }
                RewriteElements(view, "JSLink", target.RenderingResourceRewrites, replacements);
                RewriteElements(view, "XslLink", target.RenderingResourceRewrites, replacements);
                definition.Value = view.ToString(SaveOptions.DisableFormatting);
            }
            RewriteProperty(document, "JSLink", target.RenderingResourceRewrites, replacements);
            RewriteProperty(document, "XslLink", target.RenderingResourceRewrites, replacements);
            var xml = document.ToString(SaveOptions.DisableFormatting);
            return new RewrittenClassicWebPart
            {
                SourceWebPartId = binding.SourceWebPartId,
                ExportXml = xml,
                ExportSha256 = MigrationDigest.ComputeSha256(xml),
                Replacements = replacements
            };
        }

        private static void SetProperty(XDocument document, string name, string value, IDictionary<string, string> replacements)
        {
            var property = FindProperty(document, name);
            if (property == null)
            {
                throw new InvalidDataException("The Web Part export has no '" + name + "' property.");
            }
            replacements["property:" + name + ":" + property.Value] = value;
            property.Value = value;
        }

        private static XElement FindProperty(XDocument document, string name)
        {
            return document.Descendants().LastOrDefault(element => string.Equals(element.Name.LocalName, "property", StringComparison.OrdinalIgnoreCase)
                && string.Equals((string)element.Attribute("name"), name, StringComparison.OrdinalIgnoreCase));
        }

        private static void SetOptionalProperty(XDocument document, string name, string value, IDictionary<string, string> replacements)
        {
            var property = FindProperty(document, name);
            if (property == null)
            {
                return;
            }
            replacements["property:" + name + ":" + property.Value] = value;
            property.Value = value;
        }

        private static void RewriteProperty(XDocument document, string name, IDictionary<string, string> rewrites, IDictionary<string, string> replacements)
        {
            var property = FindProperty(document, name);
            if (property == null)
            {
                return;
            }
            var rewritten = RewriteTokens(property.Value, rewrites);
            if (!string.Equals(rewritten, property.Value, StringComparison.Ordinal))
            {
                replacements["property:" + name + ":" + property.Value] = rewritten;
                property.Value = rewritten;
            }
        }

        private static void RewriteElements(XDocument document, string name, IDictionary<string, string> rewrites, IDictionary<string, string> replacements)
        {
            foreach (var element in document.Descendants().Where(value => string.Equals(value.Name.LocalName, name, StringComparison.OrdinalIgnoreCase)))
            {
                var rewritten = RewriteTokens(element.Value, rewrites);
                if (!string.Equals(rewritten, element.Value, StringComparison.Ordinal))
                {
                    replacements["XmlDefinition/" + name + ":" + element.Value] = rewritten;
                    element.Value = rewritten;
                }
            }
        }

        private static string RewriteTokens(string value, IDictionary<string, string> rewrites)
        {
            if (rewrites == null)
            {
                return value;
            }
            string direct;
            if (rewrites.TryGetValue((value ?? string.Empty).Trim(), out direct))
            {
                return direct;
            }
            if (string.IsNullOrEmpty(value) || value.IndexOf('|') < 0)
            {
                return value;
            }
            return string.Join("|", value.Split('|').Select(token =>
            {
                var trimmed = token.Trim();
                string replacement;
                return rewrites.TryGetValue(trimmed, out replacement) ? replacement : trimmed;
            }));
        }
    }
}
