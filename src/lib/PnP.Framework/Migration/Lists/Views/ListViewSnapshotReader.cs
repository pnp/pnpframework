using Microsoft.SharePoint.Client;
using PnP.Framework.Migration.Packaging;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;

namespace PnP.Framework.Migration.Lists.Views
{
    internal static class ListViewSnapshotReader
    {
        public static IList<ListViewSnapshot> Read(ViewCollection views, string listRootFolder)
        {
            return views.AsEnumerable().Select(view =>
            {
                var xml = view.ListViewXml ?? string.Empty;
                return new ListViewSnapshot
                {
                    Id = view.Id,
                    Title = view.Title ?? string.Empty,
                    ServerRelativeUrl = view.ServerRelativeUrl,
                    Hidden = view.Hidden,
                    DefaultView = view.DefaultView,
                    PersonalView = view.PersonalView,
                    ViewType = view.ViewType,
                    RowLimit = view.RowLimit,
                    Paged = view.Paged,
                    ViewQuery = view.ViewQuery ?? string.Empty,
                    ViewFields = view.ViewFields.ToList(),
                    ListViewXml = xml,
                    ListViewXmlSha256 = MigrationDigest.ComputeSha256(xml),
                    JsLink = NullIfEmpty(view.JSLink) ?? ReadElement(xml, "JSLink"),
                    XslLink = ReadElement(xml, "XslLink"),
                    IsPageBound = !IsBelow(view.ServerRelativeUrl, listRootFolder)
                };
            }).OrderBy(value => value.Id).ToList();
        }

        private static string ReadElement(string xml, string name)
        {
            if (string.IsNullOrWhiteSpace(xml))
            {
                return null;
            }
            try
            {
                return XDocument.Parse(xml).Descendants().FirstOrDefault(value => string.Equals(value.Name.LocalName, name, StringComparison.OrdinalIgnoreCase))?.Value;
            }
            catch (System.Xml.XmlException)
            {
                return null;
            }
        }

        private static bool IsBelow(string value, string parent)
        {
            return string.Equals(value, parent, StringComparison.OrdinalIgnoreCase)
                || value.StartsWith(parent.TrimEnd('/') + "/", StringComparison.OrdinalIgnoreCase);
        }

        private static string NullIfEmpty(string value)
        {
            return string.IsNullOrWhiteSpace(value) ? null : value.Trim();
        }
    }
}
