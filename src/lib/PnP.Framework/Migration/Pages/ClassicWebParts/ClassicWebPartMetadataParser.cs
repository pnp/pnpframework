using System;
using System.Linq;
using System.Xml.Linq;

namespace PnP.Framework.Migration.Pages.ClassicWebParts
{
    internal static class ClassicWebPartMetadataParser
    {
        public static string ReadTypeName(string exportXml)
        {
            if (string.IsNullOrWhiteSpace(exportXml))
            {
                return null;
            }

            try
            {
                return XDocument.Parse(exportXml, LoadOptions.None)
                    .Descendants()
                    .Where(element => string.Equals(element.Name.LocalName, "type", StringComparison.OrdinalIgnoreCase))
                    .Select(element => (string)element.Attribute("name"))
                    .FirstOrDefault(value => !string.IsNullOrWhiteSpace(value));
            }
            catch (System.Xml.XmlException)
            {
                return null;
            }
        }
    }
}
