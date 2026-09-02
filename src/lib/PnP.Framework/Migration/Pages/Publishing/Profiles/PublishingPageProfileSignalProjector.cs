using PnP.Framework.Migration.Pages.Fields;
using PnP.Framework.Migration.Pages.Profiles;
using PnP.Framework.Migration.Pages.Publishing.Layouts;
using System;
using System.Collections.Generic;
using System.Linq;

namespace PnP.Framework.Migration.Pages.Publishing.Profiles
{
    public static class PublishingPageProfileSignalProjector
    {
        public static IList<PageProfileSignal> Project(
            PageIdentity source,
            PublishingPageLayoutSnapshot layout,
            IEnumerable<PageFieldValueSnapshot> fields)
        {
            var result = new List<PageProfileSignal>();
            var contentTypeId = source?.ContentTypeId;
            if (!string.IsNullOrWhiteSpace(contentTypeId)
                && contentTypeId.StartsWith(BuiltInContentTypeId.EnterpriseWikiPage, StringComparison.OrdinalIgnoreCase))
            {
                result.Add(Signal(
                    PageProfileIds.EnterpriseWiki,
                    PageProfileSignalKind.ContentTypeLineage,
                    contentTypeId,
                    "The source Content Type descends from Enterprise Wiki Page."));
            }

            if (!string.IsNullOrWhiteSpace(contentTypeId)
                && contentTypeId.StartsWith(BuiltInContentTypeId.ProjectPage, StringComparison.OrdinalIgnoreCase))
            {
                result.Add(Signal(
                    PageProfileIds.ProjectPage,
                    PageProfileSignalKind.ContentTypeLineage,
                    contentTypeId,
                    "The source Content Type descends from Project Page."));
            }

            AddLayoutSignals(layout?.FileName, result);
            var fieldNames = new HashSet<string>(
                (fields ?? Array.Empty<PageFieldValueSnapshot>())
                    .Where(value => value != null)
                    .Select(value => value.InternalName),
                StringComparer.OrdinalIgnoreCase);
            if (fieldNames.Contains("Wiki_x0020_Page_x0020_Categories"))
            {
                result.Add(Signal(
                    PageProfileIds.EnterpriseWiki,
                    PageProfileSignalKind.Field,
                    "Wiki_x0020_Page_x0020_Categories",
                    "The Pages item exposes the Enterprise Wiki Category field."));
            }

            foreach (var projectField in new[] { "TaskStatus", "WebPage", "PublishingContact" })
            {
                if (fieldNames.Contains(projectField))
                {
                    result.Add(Signal(
                        PageProfileIds.ProjectPage,
                        PageProfileSignalKind.Field,
                        projectField,
                        "The Pages item exposes a Project Page profile field."));
                }
            }

            return result
                .OrderBy(value => value.ProfileId, StringComparer.Ordinal)
                .ThenBy(value => value.Kind)
                .ThenBy(value => value.Subject, StringComparer.OrdinalIgnoreCase)
                .ToList();
        }

        private static void AddLayoutSignals(string fileName, ICollection<PageProfileSignal> result)
        {
            if (string.Equals(fileName, "EnterpriseWiki.aspx", StringComparison.OrdinalIgnoreCase))
            {
                result.Add(Signal(
                    PageProfileIds.EnterpriseWiki,
                    PageProfileSignalKind.Layout,
                    fileName,
                    "The selected Page Layout is EnterpriseWiki.aspx."));
            }

            if (string.Equals(fileName, "ProjectPage.aspx", StringComparison.OrdinalIgnoreCase))
            {
                result.Add(Signal(
                    PageProfileIds.ProjectPage,
                    PageProfileSignalKind.Layout,
                    fileName,
                    "The selected Page Layout is ProjectPage.aspx."));
            }
        }

        private static PageProfileSignal Signal(
            string profileId,
            PageProfileSignalKind kind,
            string subject,
            string evidence)
        {
            return new PageProfileSignal
            {
                ProfileId = profileId,
                Kind = kind,
                Subject = subject,
                Evidence = evidence
            };
        }
    }
}
