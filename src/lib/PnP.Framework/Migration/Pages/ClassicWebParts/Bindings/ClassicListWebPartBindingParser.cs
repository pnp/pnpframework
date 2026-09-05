using PnP.Framework.Migration.Diagnostics;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;

namespace PnP.Framework.Migration.Pages.ClassicWebParts.Bindings
{
    public static class ClassicListWebPartBindingParser
    {
        public static ClassicListWebPartBindingParseResult Parse(
            ClassicWebPartSnapshot webPart,
            Guid sourcePageWebId,
            string sourcePageWebUrl,
            string sourcePageServerRelativeUrl)
        {
            if (webPart == null)
            {
                throw new ArgumentNullException(nameof(webPart));
            }

            var issues = new List<MigrationIssue>();
            XDocument document;
            try
            {
                document = XDocument.Parse(webPart.ExportXml, LoadOptions.PreserveWhitespace);
            }
            catch (System.Xml.XmlException exception)
            {
                AddBlocker(issues, webPart.Id, "ListBindingUnavailable", "The list-bound Web Part export XML is malformed: " + exception.Message);
                return new ClassicListWebPartBindingParseResult { Issues = issues };
            }

            var properties = document.Descendants()
                .Where(element => string.Equals(element.Name.LocalName, "property", StringComparison.OrdinalIgnoreCase))
                .Where(element => !string.IsNullOrWhiteSpace((string)element.Attribute("name")))
                .GroupBy(element => (string)element.Attribute("name"), StringComparer.OrdinalIgnoreCase)
                .ToDictionary(group => group.Key, group => group.Last().Value.Trim(), StringComparer.OrdinalIgnoreCase);
            string listIdValue;
            properties.TryGetValue("ListId", out listIdValue);
            string listNameValue;
            properties.TryGetValue("ListName", out listNameValue);
            var listId = ParseGuid(listIdValue) ?? ParseGuid(listNameValue);
            if (!listId.HasValue)
            {
                AddBlocker(issues, webPart.Id, "ListBindingUnavailable", "A list-bound Web Part has no parseable ListId/ListName GUID.");
            }

            string webIdValue;
            properties.TryGetValue("WebId", out webIdValue);
            var declaredWebId = ParseGuid(webIdValue);
            var sourceListWebId = !declaredWebId.HasValue || declaredWebId.Value == Guid.Empty ? sourcePageWebId : declaredWebId.Value;
            string xmlDefinition;
            properties.TryGetValue("XmlDefinition", out xmlDefinition);
            xmlDefinition = xmlDefinition ?? string.Empty;
            string viewGuid;
            properties.TryGetValue("ViewGuid", out viewGuid);
            var viewId = ParseGuid(viewGuid);
            string jsLink;
            properties.TryGetValue("JSLink", out jsLink);
            string xslLink;
            properties.TryGetValue("XslLink", out xslLink);
            if (!string.IsNullOrWhiteSpace(xmlDefinition))
            {
                try
                {
                    var view = XDocument.Parse(xmlDefinition, LoadOptions.PreserveWhitespace).Root;
                    viewId = viewId ?? ParseGuid(view == null ? null : (string)view.Attribute("Name"));
                    jsLink = FirstNonempty(jsLink, ReadElement(view, "JSLink"));
                    xslLink = FirstNonempty(xslLink, ReadElement(view, "XslLink"));
                }
                catch (System.Xml.XmlException exception)
                {
                    AddBlocker(issues, webPart.Id, "ViewMappingUnavailable", "The embedded view definition is malformed: " + exception.Message);
                }
            }
            else
            {
                AddBlocker(issues, webPart.Id, "ViewMappingUnavailable", "The list-bound Web Part has no captured XmlDefinition/CAML view.");
            }

            if (!listId.HasValue || issues.Any(value => value.Severity == MigrationIssueSeverity.Blocker))
            {
                return new ClassicListWebPartBindingParseResult { Issues = issues };
            }

            string titleUrl;
            properties.TryGetValue("TitleUrl", out titleUrl);
            return new ClassicListWebPartBindingParseResult
            {
                Binding = new ClassicListWebPartBindingSnapshot
                {
                    SourceWebPartId = webPart.Id,
                    TypeName = webPart.TypeName ?? ClassicWebPartMetadataParser.ReadTypeName(webPart.ExportXml),
                    Title = webPart.Title,
                    SourcePageWebId = sourcePageWebId,
                    SourcePageWebUrl = sourcePageWebUrl,
                    SourcePageServerRelativeUrl = sourcePageServerRelativeUrl,
                    SourceListWebId = sourceListWebId,
                    SourceListId = listId.Value,
                    SourceViewId = viewId,
                    SourceListServerRelativeUrl = ServerRelativePath(titleUrl, sourcePageWebUrl),
                    SourceTitleUrl = titleUrl,
                    XmlDefinition = xmlDefinition,
                    JsLink = NullIfEmpty(jsLink),
                    XslLink = NullIfEmpty(xslLink),
                    SourceExportSha256 = webPart.ExportSha256,
                    SourceExportXml = webPart.ExportXml
                },
                Issues = issues
            };
        }

        public static bool IsListBound(ClassicWebPartSnapshot webPart)
        {
            if (webPart == null || string.IsNullOrWhiteSpace(webPart.ExportXml))
            {
                return false;
            }
            try
            {
                return XDocument.Parse(webPart.ExportXml).Descendants()
                    .Where(element => string.Equals(element.Name.LocalName, "property", StringComparison.OrdinalIgnoreCase))
                    .Any(element => string.Equals((string)element.Attribute("name"), "ListId", StringComparison.OrdinalIgnoreCase)
                        || string.Equals((string)element.Attribute("name"), "ListName", StringComparison.OrdinalIgnoreCase));
            }
            catch (System.Xml.XmlException)
            {
                return false;
            }
        }

        private static string ReadElement(XElement root, string name)
        {
            return root == null ? null : root.Elements().FirstOrDefault(value => string.Equals(value.Name.LocalName, name, StringComparison.OrdinalIgnoreCase))?.Value;
        }

        private static Guid? ParseGuid(string value)
        {
            Guid result;
            return Guid.TryParse((value ?? string.Empty).Trim().Trim('{', '}'), out result) ? result : (Guid?)null;
        }

        private static string ServerRelativePath(string value, string sourceWebUrl)
        {
            if (string.IsNullOrWhiteSpace(value))
            {
                return null;
            }
            Uri absolute;
            if (Uri.TryCreate(value, UriKind.Absolute, out absolute))
            {
                return Uri.UnescapeDataString(absolute.AbsolutePath);
            }
            if (value.StartsWith("/", StringComparison.Ordinal))
            {
                return value;
            }
            return new Uri(sourceWebUrl).AbsolutePath.TrimEnd('/') + "/" + value.TrimStart('/');
        }

        private static string FirstNonempty(string first, string second)
        {
            return !string.IsNullOrWhiteSpace(first) ? first : NullIfEmpty(second);
        }

        private static string NullIfEmpty(string value)
        {
            return string.IsNullOrWhiteSpace(value) ? null : value.Trim();
        }

        private static void AddBlocker(ICollection<MigrationIssue> issues, Guid webPartId, string code, string message)
        {
            issues.Add(new MigrationIssue
            {
                Code = code,
                Severity = MigrationIssueSeverity.Blocker,
                Subject = "webpart:" + webPartId.ToString("D"),
                Ingredient = "ClassicWebPart.ListBinding",
                Message = message
            });
        }
    }
}
